Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name									|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_MailUtil_SendEmail							|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to send an email
'002. 	Fn_MailUtil_EmailBodyHTML						|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to generate email HTML body
'003. 	Fn_MailUtil_GetBatchExecutionResultDetail		|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to get batch execution result details
'004. 	Fn_MailUtil_GetBatchExecutionStartDateAndTime	|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to get batch execution start date and time
'005. 	Fn_MailUtil_GetBatchExecutionDuration			|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to get batch execution duration
'006. 	Fn_MailUtil_TestSetOperations					|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to perform operation on Test set
'007. 	Fn_MailUtil_BatchExecutionResultCleanup			|	Sandeep.Navghane@sqs.com	|	15-Mar-2015	|	Function used to cleanup batch execution result
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_SendEmail
'
'Function Description	 :	Function used to send an email
'
'Function Parameters	 :  1.sSendFrom			: Mail from email id's
'							2.sSendTo			: Mail recipient ( TO ) email id's separated by semicolon (;)
'							3.sSendCC			: Mail recipient ( TO ) email id's separated by semicolon (;)
'							4.sSubject			: Mail subject
'							5.sHTMLBody			: Mail body contents
'							6.sAttachmentPath	: Mail attahcments path
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Valid mail settings and email address
'
'Function Usage		     :  bReturn = Fn_MailUtil_SendEmail("sandeep.navghane@sqs.com","Kundan.Kudale@SQS.com","","This is an Automated email","Testing","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MailUtil_SendEmail(sSendFrom,sSendTo,sSendCC,sSubject,sHTMLBody,sAttachmentPath)
	'Declaring variables
	Dim aAttachmentPath
	Dim iCounter
	Dim objMessage
	
	Fn_MailUtil_SendEmail=False
	
	'Creating object of CDO
	Set objMessage = CreateObject("CDO.Message")
	
	'This section provides the configuration information for the remote SMTP server.
	'Normally you will only change the server name or IP.
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	
	'Name or IP of Remote SMTP Server
	'please pass valid smptserver name here
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail1.sqs.com"
	
	'Server port (typically 25)
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	objMessage.Configuration.Fields.Update
	objMessage.Subject = sSubject
	objMessage.From = sSendFrom
	objMessage.To = sSendTo
	objMessage.CC = sSendCC
    objMessage.HTMLBody = sHTMLBody
    'objMessage.AddAttachment sAttachmentPath
    If sAttachmentPath<>"" Then
    	aAttachmentPath=Split(sAttachmentPath,";")
    	For iCounter=0 To Ubound(aAttachmentPath)
    		objMessage.AddAttachment aAttachmentPath(iCounter)		    		
    	Next
    End If
	objMessage.Send
	Set objMessage = Nothing
	Fn_MailUtil_SendEmail = True
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_EmailBodyHTML
'
'Function Description	 :	Function used to generate email HTML body
'
'Function Parameters	 :  1.sExcelPath	: Batch execution result excel file path
'
'Function Return Value	 : 	Email HTML body 
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_MailUtil_EmailBodyHTML("C:\Mainline\Reports\BatchExecutionDetails.xlsx")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MailUtil_EmailBodyHTML(sExcelPath)
	'Declaring variables
	Dim sBatchOwner,sBatchName,sBatchResult,sBatchDate,sBatchDuration,sQCBatchLogFolder
	Dim sBatchArea,sMailBody,sSolidWorksURL,sTeamcenterURL,sTestArea,sComputerName
	Dim sBatchTotalTestCases,sBatchPassTestCases,sBatchFailTestCases,sBrowserUsed
	Dim sBatchStartDateAndTime,sBatchEndDateAndTime
	Dim objWScriptShell,objFile,objFSO	
	Dim bReturn
	
	'Getting Batch Owner name
	If lcase(Environment("UserName")) = "" Then
		sBatchOwner = Fn_CommonUtil_LocalMachineOperations("getcurrentloginusername","")
	Else
		sBatchOwner = Environment("UserName")
	End If
	
	'Getting machine name
	Set objWScriptShell =CreateObject( "WScript.Shell" )
	sComputerName = objWScriptShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )
	Set objWScriptShell =Nothing
	
	'Batch result
	bReturn = Fn_MailUtil_GetBatchExecutionResultDetail(sExcelPath,"FailExist", 1)
	if bReturn = "True" Then
		'sBatchResult = "FAIL" 
		sBatchResult = "<tr><td><font face=""Arial"" SIZE=2>Test Set Execution Result</td>" &_
					   "<td BGCOLOR=""#FF7373"" align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;FAIL</B></td></tr>"
	else
		'sBatchResult = "PASS"
		sBatchResult = "<tr><td><font face=""Arial"" SIZE=2>Test Set Execution Result</td>" &_
					   "<td BGCOLOR=""#7DDBA9"" align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;PASS</B></td></tr>"
	end if	
	'Batch Name
	sBatchName=""
	'If test management tool is ALM\QC then uncomment below line
	'sBatchName = Fn_MailUtil_TestSetOperations("GetTestSetName")
	
	'Batch date
	sBatchStartDateAndTime = Fn_MailUtil_GetBatchExecutionStartDateAndTime(sExcelPath)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(sExcelPath)
	sBatchEndDateAndTime = objFile.DateLastModified
	Set objFile =Nothing
	Set objFSO = Nothing
	'Batch duration
	sBatchDuration = Fn_MailUtil_GetBatchExecutionDuration(sExcelPath)
	'Batch log folder
	
	sQCBatchLogFolder=""
	'If test management tool is ALM\QC then uncomment below line
	'sQCBatchLogFolder = Fn_MailUtil_TestSetOperations("GetTestSetPath")
	
	'Testing area
	sTestArea = "Regression Testing"
	'Cleanup batch result 
	bReturn = Fn_MailUtil_BatchExecutionResultCleanup(sExcelPath, 1)
	' Total testcases in batch
	sBatchTotalTestCases = Fn_MailUtil_GetBatchExecutionResultDetail(sExcelPath, "TotalTestCaseCount", 1)
	'Pass test cases in batch
	sBatchPassTestCases = Fn_MailUtil_GetBatchExecutionResultDetail(sExcelPath, "PassCount", 1)
	'Browser Used	
	If LCase(Environment.Value("BrowserName"))="ie" Then
		sBrowserUsed = "Internet Explorer"
	Else
		sBrowserUsed =Environment.Value("BrowserName")
	End If	
	'Fail test cases in batch
	sBatchFailTestCases = sBatchTotalTestCases - sBatchPassTestCases
		
	sMailBody = ""
	sMailBody = "<HTML><HEAD></HEAD>" & _
				"<BODY>" & _
				"<table border=1 cellpadding=1 cellspacing=1 bordercolor=""gray"" width=100%>" & _
				"<tr><th colspan=2 bgcolor=""#A0CFEC""><font face=""Arial"" size=4>Automation Test Result</th></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Test Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Execution Test Area</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sTestArea & " </td></B></tr>" & _
				sBatchResult & _
				"<tr><td><font face=""Arial"" SIZE=2>Total Test Cases in Batch</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchTotalTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Number of Passed Test Cases</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchPassTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Number of Failed Test Cases</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2><B>&nbsp;" & sBatchFailTestCases & "</td></B></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Folder Path</td><td align=""left"" valign=""center"">" & _
				"<font face=""Arial"" SIZE=2><B>"& sQCBatchLogFolder &"</B></a></td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Machine Local Reports Folder Path</td><td align=""left"" valign=""center"">" & _
				"<font face=""Arial"" SIZE=2><B>"& Environment.Value("BatchFolderName") &"</B></a></td></tr>" & _					
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Batch Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Name</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchName & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Execution Start Date and Time</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchStartDateAndTime & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Execution End Date and Time</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchEndDateAndTime & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Executed By</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBatchOwner & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Executed On Machine</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sComputerName & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Test Set Execution Duration</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;<font face=""Arial"" SIZE=2>" & sBatchDuration & "</td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Teamcenter Server Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Release</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment.Value("TcRelease") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Server Setup Type</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;4 Tier</td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Teamcenter WebClient Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter WebClient URL</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment.Value("TcWebURL") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Browser Used</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & sBrowserUsed & "</td></tr>" & _
				"<tr><th colspan=2 align=""left"" bgcolor=""#A0CFEC""><font face=""Arial"" SIZE=3>Teamcenter Rich Client Details</th></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Rich Client URL</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment.Value("RACAppExecutable") & "</td></tr>" & _
				"<tr><td><font face=""Arial"" SIZE=2>Teamcenter Rich Client Database</td>" & _
				"<td align=""left"" valign=""center""><font face=""Arial"" SIZE=2>&nbsp;" & Environment.Value("Site1") & "</td></tr>" & _					
				"</table>" & _
				"<p><font face=""Arial"" SIZE=2>Thanks</p>" & _
				"<p><font face=""Arial"" SIZE=2>SQS Automation Team" & _
				"<p><font face=""Arial"" SIZE=1><br><br><br>PS: This is a automated mail, please do not reply</p>" & _
				"</BODY></HTML>"

	Fn_MailUtil_EmailBodyHTML = sMailBody
	
	'Creating batch execution summary report html file
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.CreateTextFile(Environment.Value("BatchFolderName") & "\BatchRunSummaryReport.html")
	objFile.Write sMailBody
	Set objFile =Nothing
	Set objFSO = Nothing
	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_GetBatchExecutionResultDetail
'
'Function Description	 :	Function used to get batch execution result details
'
'Function Parameters	 :  1.sExcelPath		: Batch execution result file path
'							2.sActionType		: Action type
'							3.iSheetNumber		: Excel file sheet number
'
'Function Return Value	 : 	True \ False \ Counts
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	batch execution result file should be exist
'
'Function Usage		     :  bReturn = Fn_MailUtil_GetBatchExecutionResultDetail("C:\Mainline\Reports\BatchExecutionDetails.xlsx", "PassCount", 1)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_MailUtil_GetBatchExecutionResultDetail(sExcelPath,sActionType,iSheetNumber)
	'Declaring variables
	Const xlCellTypeLastCell = 11
	Dim objFSO,objExcel,objWorkbook,objWorkSheet,objRange
	Dim iUsedRange,iColumnNumber,iPassCount,iFailCount,iTotalCount
	
	Fn_MailUtil_GetBatchExecutionResultDetail=False
	
	'Getting batch execution result file path
    If sExcelPath="" Then
    	sExcelPath=Environment.Value("BatchFolderName") & "\BatchRunDetails.xlsx"
    End If
	
	If iSheetNumber = "" Then
		iSheetNumber = 1
	End If 
			
	'Getting Result coumn number
	iColumnNumber=-1
	iColumnNumber = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sExcelPath,iSheetNumber,"","Result",True)
	iColumnNumber = Cint(Split(iColumnNumber,":")(1))
	
	If iColumnNumber=-1 Then
		Exit Function
	End IF
	
	iColumnNumber = Chr(64 + iColumnNumber)
	
	'Creating file system object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(sExcelPath) Then
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		objExcel.AlertBeforeOverwriting = False
		objExcel.DisplayAlerts = False
	
		Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)	
		
		Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets(iSheetNumber)
		Set objRange = objWorkSheet.UsedRange
		objRange.SpecialCells(xlCellTypeLastCell).Activate				
		iUsedRange = iColumnNumber & "2:" & iColumnNumber & objExcel.ActiveCell.Row
			
		If  LCase(sActionType) = LCase("FailExist") Then
			objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & iUsedRange & "," & Chr(34) & "FAIL" & Chr(34) & ")"
			iFailCount = objExcel.Cells(10000, 1).value
			If  iFailCount > 0 then
				Fn_MailUtil_GetBatchExecutionResultDetail = True
			Else
				iTotalCount = objExcel.ActiveCell.Row - 1
				objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & iUsedRange & "," & Chr(34) & "PASS" & Chr(34) & ")"
				iPassCount = objExcel.Cells(10000, 1).value
				objExcel.Cells(10001, 1).Formula = "=COUNTIF(" & iUsedRange & "," & Chr(34) & "FAIL" & Chr(34) & ")"
				iFailCount = objExcel.Cells(10001, 1).value
				If  iPassCount + iFailCount <> iTotalCount then
					Fn_MailUtil_GetBatchExecutionResultDetail = True
				Else
					Fn_MailUtil_GetBatchExecutionResultDetail = False
					End if
				End if																																			
		Elseif LCase(sActionType) = LCase("TotalTestCaseCount") Then
			Fn_MailUtil_GetBatchExecutionResultDetail = objExcel.ActiveCell.Row - 1								
		Elseif LCase(sActionType) = LCase("PassCount") then			
			objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & iUsedRange & "," & Chr(34) & "PASS" & Chr(34) & ")"
			Fn_MailUtil_GetBatchExecutionResultDetail = objExcel.Cells(10000, 1).value				
		Elseif LCase(sActionType) = LCase("FailCount") then
			iTotalCount = objExcel.ActiveCell.Row - 1
			objExcel.Cells(10000, 1).Formula = "=COUNTIF(" & iUsedRange & "," & Chr(34) & "PASS" & Chr(34) & ")"
			iPassCount = objExcel.Cells(10000, 1).value
			objExcel.Cells(10001, 1).Formula = "=COUNTIF(" & iUsedRange & "," & Chr(34) & "FAIL" & Chr(34) & ")"
			iFailCount = objExcel.Cells(10001, 1).value
			If  iPassCount + iFailCount <> iTotalCount then
				Fn_MailUtil_GetBatchExecutionResultDetail = iTotalCount - iPassCount
			Else
				Fn_MailUtil_GetBatchExecutionResultDetail = iFailCount																	      	
			End if								
		End if														
		objExcel.Quit			
		'Releasing all objects		
		Set objRange = Nothing
		Set objWorkSheet = Nothing
		Set objWorkbook = Nothing
		Set objExcel = Nothing							
	End If
	'Releasing file system object
	Set objFSO = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_GetBatchExecutionStartDateAndTime
'
'Function Description	 :	Function used to get batch execution start date and time
'
'Function Parameters	 :  1.sExcelPath	: Batch execution result file path
'
'Function Return Value	 : 	Batch execution start date and time \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	batch execution result file should be exist
'
'Function Usage		     :  bReturn = Fn_MailUtil_GetBatchExecutionResultDetail("C:\Mainline\Reports\BatchExecutionDetails.xlsx")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MailUtil_GetBatchExecutionStartDateAndTime(sExcelPath)
	'Declaring variables
	Const xlCellTypeLastCell = 11
	Dim objFSO,objExcel,objWorkbook,objWorkSheet,objRange
	Dim sStartDateAndTime,iColumnNumber
	
	Fn_MailUtil_GetBatchExecutionStartDateAndTime=False
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists (sExcelPath) Then
		'Getting Result coumn number
		iColumnNumber=-1
		iColumnNumber = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sExcelPath,"","","Start Time",True)
		iColumnNumber = Cint(Split(iColumnNumber,":")(1))
		
		If iColumnNumber=-1 Then
			Exit Function
		End IF
		'Creating object of excel
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False
		objExcel.AlertBeforeOverwriting = False
		objExcel.DisplayAlerts = False
			
		Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
		Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets(1)
		Set objRange = objWorkSheet.UsedRange
		objRange.SpecialCells(xlCellTypeLastCell).Activate
		sStartDateAndTime=objWorkSheet.Cells(2,iColumnNumber).Value
		sStartDateAndTime=Replace(sStartDateAndTime,"st- ","")
		sStartDateAndTime=sStartDateAndTime
		
		Fn_MailUtil_GetBatchExecutionStartDateAndTime=sStartDateAndTime
		'Releasing all objects
		Set objRange = Nothing
		Set objWorkSheet = Nothing
		Set objWorkbook = Nothing
		'Exiting from excel
		objExcel.Quit		
		Set objExcel =Nothing
	End If
	'Releasing file system object
	Set objFSO =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_GetBatchExecutionDuration
'
'Function Description	 :	Function used to get batch execution duration
'
'Function Parameters	 :  1.sExcelPath	: Batch execution result file path
'
'Function Return Value	 : 	Batch execution duration \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	batch execution result file should be exist
'
'Function Usage		     :  bReturn = Fn_MailUtil_GetBatchExecutionDuration("C:\Mainline\Reports\BatchExecutionDetails.xlsx")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MailUtil_GetBatchExecutionDuration(sExcelPath)
	'Declaring variables
	Dim objFSO,objFile
	Dim sFileCreationDate,sFileLastModifiedDate,sTimeDifference
	Dim sTimeSeparator
	
	Fn_MailUtil_GetBatchExecutionDuration = False
	
	'Creating file system object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Checking existance of batch execution file
	If objFSO.FileExists (sExcelPath) Then
		'Creating file object
		Set objFile = objFSO.GetFile(sExcelPath)
		'Getting batch execution start date and time					
		sFileCreationDate = Fn_MailUtil_GetBatchExecutionStartDateAndTime(sExcelPath)
		sFileCreationDate=CDate(sFileCreationDate)
		'Getting batch execution file last modified date and time					
		sFileLastModifiedDate = objFile.DateLastModified
		
		sTimeDifference = Formatnumber((DateDiff("s", sFileCreationDate, sFileLastModifiedDate)/3600), 2, 0, -1)
		
		If inStr(1,CStr(sTimeDifference),",") Then			
			sTimeSeparator=","			
		Else
			sTimeSeparator="."
		End If
		
		If sTimeDifference => 1 Then
			sTimeDifference=Split(sTimeDifference,sTimeSeparator)
			Fn_MailUtil_GetBatchExecutionDuration = sTimeDifference(0) & " hours, " & Formatnumber((sTimeDifference(1) * 0.6), 0, 0, -1) & " mins"			
		Else
			sTimeDifference=Split(sTimeDifference,sTimeSeparator)
			Fn_MailUtil_GetBatchExecutionDuration = Formatnumber((sTimeDifference(1) * 0.6), 0, 0, -1) & " mins"			
		End If
		Set objFile = nothing
	End if
	'Releasing objects
	Set objFSO = nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_TestSetOperations
'
'Function Description	 :	Function used to perform operation on Test set
'
'Function Parameters	 :  1.sAction	: Action to perform
'
'Function Return Value	 : 	Test set name \ Test set path \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_MailUtil_TestSetOperations("GetTestSetName")
'Function Usage		     :  bReturn = Fn_MailUtil_TestSetOperations("GetTestSetPath")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MailUtil_TestSetOperations(sAction)
	'Declaring varaibles
	Dim objQCUtil,objTestSetFolder	
	Dim aTemp
	
	Fn_MailUtil_TestSetOperations=False

	If Lcase(Cstr(Environment.Value("QCStepLog")))="true" Then
		'Creating current test set object
		Set objQCUtil=QCUtil.CurrentTestSet	
	End If

	Select Case sAction
		Case "GetTestSetName"
			If Lcase(Cstr(Environment.Value("QCStepLog")))="true" Then
				Fn_MailUtil_TestSetOperations=Cstr(objQCUtil.Name)
			Else
				aTemp=Split(Environment.Value("BatchFolderName"),"\")
				Fn_MailUtil_TestSetOperations=aTemp(Ubound(aTemp))
			End IF
		Case "GetTestSetPath"
			If Lcase(Cstr(Environment.Value("QCStepLog")))="true" Then
				Set objTestSetFolder=objQCUtil.TestSetFolder
				Fn_MailUtil_TestSetOperations=Cstr(objTestSetFolder.Path)
				Set objTestSetFolder=Nothing
			Else
				Fn_MailUtil_TestSetOperations=Environment.Value("BatchFolderName")
			End IF
	End Select	
	'Releasing current test set object
	Set objQCUtil=Nothing	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MailUtil_BatchExecutionResultCleanup
'
'Function Description	 :	Function used to cleanup batch execution result
'
'Function Parameters	 :  1.sExcelPath	:  Batch execution file path
'							2.iSheetNumber		: Excel file sheet number
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_MailUtil_BatchExecutionResultCleanup("C:\AUTMainline\Reports\BatchExecutionDetails.xlsx","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  15-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MailUtil_BatchExecutionResultCleanup(sExcelPath,iSheetNumber)
	'Declaring varaibles
	Const xlCellTypeLastCell = 11
	Const xlCenter = -4108	
	Dim objExcel,objWorkbook,objWorksheet,objRange
	Dim iRowNumber,iColumnNumber
	
	Fn_MailUtil_BatchExecutionResultCleanup = False	
	If iSheetNumber = "" Then
		 iSheetNumber = 1
	End If 	
	
	'Getting Result coumn number
	iColumnNumber=-1
	iColumnNumber = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition",sExcelPath,iSheetNumber,"","Result",True)
	iColumnNumber = Cint(Split(iColumnNumber,":")(1))
	
	If iColumnNumber=-1 Then
		Exit Function
	End IF
	'Creating excel file operation
	Set objExcel = CreateObject("Excel.Application")
	Set objWorkbook = objExcel.Workbooks.Open(sExcelPath)
	objExcel.Visible = false
	objExcel.DisplayAlerts = false
	
	Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
	objWorksheet.Activate
	objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
	'Creating object of used range
	Set objRange = objWorkSheet.UsedRange
	
	For iRowNumber = 2 To objExcel.ActiveCell.Row 
		If LCase(objExcel.Cells(iRowNumber, iColumnNumber).Value) <> LCase("PASS") AND LCase(objExcel.Cells(iRowNumber, iColumnNumber).Value) <> LCase("FAIL") then
			If Instr(LCase(objExcel.Cells(iRowNumber, iColumnNumber).Value), LCase("PASS")) <> 0 then
				objExcel.Cells(iRowNumber, iColumnNumber).Value = "PASS"
				objExcel.Cells(iRowNumber, iColumnNumber).Interior.ColorIndex = 35
			Else
				objExcel.Cells(iRowNumber, iColumnNumber).Value = "FAIL"
				objExcel.Cells(iRowNumber, iColumnNumber).Interior.Color = RGB(255,128,128)						
			End if		
		Else
			If LCase(objExcel.Cells(iRowNumber, iColumnNumber).Value) = LCase("PASS") Then
				objExcel.Cells(iRowNumber, iColumnNumber).Interior.ColorIndex = 35
			Else
				objExcel.Cells(iRowNumber, iColumnNumber).Interior.Color = RGB(255,128,128)
			End If
		End If
		objExcel.Cells(iRowNumber, iColumnNumber).Font.Bold = True
		objExcel.Cells(iRowNumber, iColumnNumber).HorizontalAlignment = xlCenter
		objExcel.Cells(iRowNumber, iColumnNumber).VerticalAlignment = xlCenter			
	Next	
	objRange.Range("A1").Activate
	'Saving and quiting from excel
	objWorkbook.Save
	objExcel.Quit
	'Releasing all object
	Set objWorksheet = Nothing
	Set objWorkbook = Nothing
	Set objExcel = Nothing			
    Fn_MailUtil_BatchExecutionResultCleanup = True
End Function
