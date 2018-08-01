'! @Name 			FMW_Setup_LocalFileAndFolderOperations
'! @Details 		To perform operations on local file and folders
'! @InputParam1 	sAction 			: Action name
'! @InputParam2 	sFileOrFolderPath 	: File path or folder path
'! @InputParam3 	sContent 			: File contents
'! @InputParam4 	sValue 				: New values or folder destination path
'! @InputParam5 	sShareName 			: Desktop sharing Name for Share
'! @InputParam6 	sCompName			: Name of computer/IP Address
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			29 Mar 2016
'! @Version 		1.0
'! @Example  		LoadAndRunAction "FMW_Setup\FMW_Setup_LocalFileAndFolderOperations","FMW_Setup_LocalFileAndFolderOperations",oneIteration,"fileexist","C:\AUT\Test.txt","","","",""
'! @Example  		LoadAndRunAction "FMW_Setup\FMW_Setup_LocalFileAndFolderOperations","FMW_Setup_LocalFileAndFolderOperations",oneIteration,"folderexist","C:\TestFolder","","","",""
'! @Example  		LoadAndRunAction "FMW_Setup\FMW_Setup_LocalFileAndFolderOperations","FMW_Setup_LocalFileAndFolderOperations",oneIteration,"createexcelfilefromtextfileatsamelocation","C:\TestFolder\abc.txt",vbTab,"","",""
'! @Example  		LoadAndRunAction "FMW_Setup\FMW_Setup_LocalFileAndFolderOperations","FMW_Setup_LocalFileAndFolderOperations",oneIteration,"verifytextinpdffile","C:\TestFolder\abc.txt","Approved~Approver: a400237~Confidential","ClosePDF","",""
'! @Example  		LoadAndRunAction "FMW_Setup\FMW_Setup_LocalFileAndFolderOperations","FMW_Setup_LocalFileAndFolderOperations",oneIteration,"verifyimageinpdffile","C:\TestFolder\abc.txt","PDF2DCircle","ClosePDF","",""

Option Explicit
Err.Clear


'Variables Declaration
Dim sAction, sFileOrFolderPath,sContent,sValue,sShareName,sCompName, sPDFText,sExcelText
Dim aFileOrFolderPath,aFolderPath,sFolderPath, sExcelFilePath, aTextToVerify,sDestinationPath
Dim iCounter,iFileSize
Dim objClipboard, objShell, objInsightObject
Dim sTempValue

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="NA"

'Get action parameters in local variables
sAction = Parameter("sAction")
sFileOrFolderPath = Parameter("sFileOrFolderPath")
sContent = Parameter("sContent")
sValue = Parameter("sValue")
sShareName = Parameter("sShareName")
sCompName = Parameter("sCompName")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "FMW_Setup_LocalFileAndFolderOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

If sFileOrFolderPath<>"" Then	
	aFolderPath=Split(sFileOrFolderPath,"\")
	If Instr(1,aFolderPath(0),"<<")>0 and Instr(1,aFolderPath(0),">>")>0 Then
		sFolderPath=Replace(aFolderPath(0),"<<","")
		sFolderPath=Replace(sFolderPath,">>","")
		sFolderPath=Fn_Setup_GetAutomationFolderPath(sFolderPath)
		sFileOrFolderPath=Replace(sFileOrFolderPath,aFolderPath(0),sFolderPath)
	End If
End If

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "fileexist"
		If Fn_FSOUtil_FileOperations("fileexist",sFileOrFolderPath,"","")=True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified a file is exist at location [ " & Cstr(sFileOrFolderPath) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as a file does not exist at location [ " & Cstr(sFileOrFolderPath) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "multiplefileexist"
		aFileOrFolderPath=Split(sFileOrFolderPath,"~")
		For iCounter=0 to Ubound(aFileOrFolderPath)
			If Fn_FSOUtil_FileOperations("fileexist",aFileOrFolderPath(iCounter),"","")=True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified a file is exist at location [ " & Cstr(aFileOrFolderPath(iCounter)) & " ]","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as a file does not exist at location [ " & Cstr(aFileOrFolderPath(iCounter)) & " ]","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "verifyfilesizeismorethanzero"
		iFileSize=Fn_FSOUtil_FileOperations("getfilesize",sFileOrFolderPath,"","")
		If cdbl(iFileSize)>0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified size of a file [ " & Cstr(sFileOrFolderPath) & " ] is more than 0 kb,current size of file is [ " & Cstr(iFileSize) & " ] kb","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as size of a file [ " & Cstr(sFileOrFolderPath) & " ]  is 0 kb","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "folderexist"
		If Fn_FSOUtil_FolderOperationss("exist",sFileOrFolderPath,"","","")=True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified a folder is exist at location [ " & Cstr(sFileOrFolderPath) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as a folder does not exist at location [ " & Cstr(sFileOrFolderPath) & " ]","","","","","")
			Call Fn_ExitTest()
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "verifyfilecontent"
		If Fn_FSOUtil_FileOperations("verifytext",sFileOrFolderPath,sContent,"")=True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified file [ " & Cstr(sFileOrFolderPath) & " ] contains [ " & Cstr(sContent) & " ] values","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as a file does not exist at location [ " & Cstr(sFileOrFolderPath) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "verifynotepadfilecontent"
		If Window("wwnd_Notepad").Exist Then			
			If Instr(1,Window("wwnd_Notepad").WinEditor("wedt_TextArea").GetROProperty("text"),sContent) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified currently open notepad file contains [ " & Cstr(sContent) & " ] value","","","","DONOTSYNC","")
			Else
				Window("wwnd_Notepad").Close
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification currently open notepad file does not contains [ " & Cstr(sContent) & " ] value","","","","","")				
				Call Fn_ExitTest()	
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify notepad file contents as notepad file is not open","","","","","")		
		End If
		Window("wwnd_Notepad").Close
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Case to Verify folder is empty
	Case "verifyfolderisempty"
		If Fn_FSOUtil_FolderOperations("getfilecount", sFileOrFolderPath,"","","")=0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified folder [ " & Cstr(sFileOrFolderPath) & " ] is empty","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as folder [ " & Cstr(sFileOrFolderPath) & " ] is not empty","","","","","")
			Call Fn_ExitTest()
		End If
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Case to create an excel file from text file
	Case "createexcelfilefromtextfileatsamelocation"
	
		'Check existence of input text file
		If Fn_FSOUtil_FileOperations("fileexist",sFileOrFolderPath,"","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of input text file at location [" & sFileOrFolderPath & "] while creating an excel file from text file","","","","","")
			Call Fn_ExitTest()
		End IF
		
		'Delete output file if already present
		sExcelFilePath = Replace(sFileOrFolderPath,".txt",".xls")
		If Fn_FSOUtil_FileOperations("fileexist",sExcelFilePath,"","") = True Then
			If Fn_FSOUtil_FileOperations("deletefile",sExcelFilePath,"","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to delete already present excel file at location [" & sExcelFilePath & "] while creating an excel file from text file","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		'Call function to create excel file from text file
		If Fn_FSOUtil_CreateExcelFromTextFile(sFileOrFolderPath, sExcelFilePath, sContent) = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create excel file at location [" & sExcelFilePath & "] while creating an excel file from text file","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created excel file at location [" & sExcelFilePath & "] while creating an excel file from text file","","","","","")
		End If
		
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Case to verify text in PDF file
	Case "verifytextinpdffile"
	
	'Open the PDF file if path is passed
	If sFileOrFolderPath <> "" Then
		If Fn_FSOUtil_FileOperations("fileexist",sFileOrFolderPath,"","") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of file [ " & sFileOrFolderPath & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			SystemUtil.Run sFileOrFolderPath
		End If
	End If
	
	'Verify existence of PDF window
	If Window("wwnd_AdobeAcrobatReader").Exist Then
		
	Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of Adobe Acrobat Reader window","","","","","")
		Call Fn_ExitTest()
	End If
	
	'Click on the pdf page
	Window("wwnd_AdobeAcrobatReader").WinObject("wobj_PDFPage").Click 10,10
	wait(GBL_MIN_MICRO_TIMEOUT)
	
	'Create object of Shell scripting
	Set objShell = CreateObject("WScript.Shell")
	
	'Create object of mercury clipboard
	Set objClipboard = CreateObject("Mercury.Clipboard")
	
	'Select all and copy the pdf content
    objShell.SendKeys "^(a)"
    wait(GBL_MIN_MICRO_TIMEOUT)
    objShell.SendKeys "^(c)"
    
    'Get data from clipboard in variable
    sPDFText = objClipboard.GetText
    
    'Verify each text value is present in PDF
    aTextToVerify = Split(sContent, "~")
	For iCounter = 0 To Ubound(aTextToVerify) Step 1
		If Instr(sPDFText, aTextToVerify(iCounter)) > 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification passed as [ " & Cstr(aTextToVerify(iCounter)) & " ] was found in pdf text [" & sPDFText & "]","","","","","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aTextToVerify(iCounter)) & " ] was not found in pdf text [" & sPDFText & "]","","","","","")
			Call Fn_ExitTest()
		End If
	Next
	
	'Close the PDF file if required
	If sValue = "ClosePDF" Then
		Window("wwnd_AdobeAcrobatReader").Close
	End If
	
	'Remove objects from memory
	Set objShell = Nothing
	Set objClipboard = Nothing
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "verifynotepadfilecontentext"
		If Window("wwnd_Notepad").Exist Then
			sTempValue=Window("wwnd_Notepad").WinEditor("wedt_TextArea").GetROProperty("text")
			If sValue="" Then
				sTempValue=Replace(sTempValue," ","")
				sTempValue=Replace(sTempValue,"-","")
				sContent=Replace(sTempValue," ","")
				sContent=Replace(sTempValue,"-","")
			End If
			If Trim(sTempValue)=Trim(sContent) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified currently open notepad file contains [ " & Cstr(sContent) & " ] value","","","","DONOTSYNC","")
			Else
				Window("wwnd_Notepad").Close
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification currently open notepad file does not contains [ " & Cstr(sContent) & " ] value","","","","","")				
				Call Fn_ExitTest()	
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify notepad file contents as notepad file is not open","","","","","")		
		End If
		Window("wwnd_Notepad").Close
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case 
	Case "copyopenexcelfromtempandpasteundertestdata"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "FMW_Setup_LocalFileAndFolderOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
		If Window("wwnd_Excel").Exist Then
			sTempValue=Window("wwnd_Excel").GetROProperty("text")
			sTempValue=Split(sTempValue," - ")
			sFolderPath=Fn_Setup_GetAutomationFolderPath("TestData")
			If Fn_FSOUtil_FileOperations("copyfile","C:\Temp\" & Trim(sTempValue(0)),"",sFolderPath & "\")=False Then
				Window("wwnd_Excel").Close
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to copy excel file from location [ C:\Temp\" & Cstr(Trim(sTempValue(0))) & " ] and paste under folder [ " & Cstr(sFolderPath) & " ]","","","","","")				
				Call Fn_ExitTest()
			End If
			DataTable.Value("ReusableActionWordReturnValue","Global")= sFolderPath & "\" & Cstr(Trim(sTempValue(0)))
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to copy excel file from [ C:\Temp ] folder as temporary excel does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Window("wwnd_Excel").Close
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate existance of local file
	Case "verifynotepadfilecontentnotavailable"
		If Window("wwnd_Notepad").Exist Then			
			If Instr(1,Window("wwnd_Notepad").WinEditor("wedt_TextArea").GetROProperty("text"),sContent) Then				
				Window("wwnd_Notepad").Close
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as currently open notepad file contains [ " & Cstr(sContent) & " ] value","","","","","")				
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified currently open notepad file does not contains [ " & Cstr(sContent) & " ] value","","","","DONOTSYNC","")	
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail to verify notepad file contents as notepad file is not open","","","","","")		
			Call Fn_ExitTest()
		End If
		If sValue = "DoNotClose" Then
			'Do not close notepad
		Else
			Window("wwnd_Notepad").Close
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate PDF file is open in new window
	Case "verifypdffileisopen"
		If Window("wwnd_AdobeAcrobatReader").Exist Then			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified PDF file is open in new window","","","","DONOTSYNC","")
			Window("wwnd_AdobeAcrobatReader").Close
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as PDF file does not open in new window","","","","","")
			Call Fn_ExitTest()
		End If		
		
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate Image file is open in new window
	Case "verifyimagefileisopen"
		If Window("wwnd_WindowsPhotoViewer").Exist Then			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Image file is open in new window","","","","DONOTSYNC","")
			Window("wwnd_WindowsPhotoViewer").Close
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as Image file does not open in new window","","","","","")
			Call Fn_ExitTest()
		End If		
	
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify text in PDF file
	Case "verifyimageinpdffile"
	
		'Open the PDF file if path is passed
		If sFileOrFolderPath <> "" Then
			If Fn_FSOUtil_FileOperations("fileexist",sFileOrFolderPath,"","") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of file [ " & sFileOrFolderPath & " ]","","","","","")
				Call Fn_ExitTest()
			Else
				SystemUtil.Run sFileOrFolderPath
			End If
		End If
		
		'Verify existence of PDF window
		If Window("wwnd_AdobeAcrobatReader").Exist Then
			Window("wwnd_AdobeAcrobatReader").Maximize
			wait 2
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of Adobe Acrobat Reader window","","","","","")
			Call Fn_ExitTest()
		End If
		
		Set objInsightObject = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_" & sContent,"")
		If objInsightObject.Exist(5) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification passed as [ " & Cstr(sContent) & " ] image was found in pdf","","","","","")
		Else
			Window("wwnd_AdobeAcrobatReader").Close
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sContent) & " ] image was not found in pdf","","","","","")
			Call Fn_ExitTest()
		End If
		
		
		'Close the PDF file if required
		If sValue = "ClosePDF" Then
			Window("wwnd_AdobeAcrobatReader").Close
		End If
		
		'Remove objects from memory
		Set objShell = Nothing
		Set objClipboard = Nothing
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate Excel file is open in new window
	Case "verifyexcelfileisopen"
		If Window("wwnd_Excel").Exist Then			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Excel file is open in new window","","","","DONOTSYNC","")
			Window("wwnd_Excel").Close
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as Excel file does not open in new window","","","","","")
			Call Fn_ExitTest()
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to vealidate Wor file is open in new window
	Case "verifywordfileisopen"
		If Window("wwnd_Word").Exist Then			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Word file is open in new window","","","","DONOTSYNC","")
			Window("wwnd_Word").Close
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as Word file does not open in new window","","","","","")
			Call Fn_ExitTest()
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify excel sheet contents
	Case "verifyexcelsheetcontents"
	 	 wait 10
	 	 If Window("wwnd_Excel").WinObject("wobj_Sheet").Exist Then		
			wait 2
			Window("wwnd_Excel").WinObject("wobj_Sheet").Click 50,30
			wait 1
			'Create object of Shell scripting
			Set objShell = CreateObject("WScript.Shell")
			'Create object of mercury clipboard
			Set objClipboard = CreateObject("Mercury.Clipboard")
			'Select all and copy the pdf content
		    objShell.SendKeys "^(a)"
		    wait(6)
		    objShell.SendKeys "^(c)"
		    wait(6)
		    'Get data from clipboard in variable
		    sExcelText = objClipboard.GetText
			Window("wwnd_Excel").WinObject("wobj_Sheet").Click 50,30
			wait 1
		    objShell.SendKeys "{ESC}"
		    wait 3
		    objClipboard.Clear
		    'Verify each text value is present in Excel
		    aTextToVerify = Split(sContent, "~")
			For iCounter = 0 To Ubound(aTextToVerify) Step 1
				If Instr(sExcelText, aTextToVerify(iCounter)) > 0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification passed as [ " & Cstr(aTextToVerify(iCounter)) & " ] was found in excel text [" & sExcelText & "]","","","","","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aTextToVerify(iCounter)) & " ] was not found in excel text [" & sExcelText & "]","","","","","")
					Call Fn_ExitTest()
				End If
			Next
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as excel sheet not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Close the excel file if required
		If sValue = "CloseExcel" Then
			Window("wwnd_Excel").Close
		End If
		
		'Remove objects from memory
		Set objShell = Nothing
		Set objClipboard = Nothing
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case 
	Case "copycatpartfromtestdataandpasteimportfolder"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "FMW_Setup_LocalFileAndFolderOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= "False"
		
	
		sFolderPath=Fn_Setup_GetAutomationFolderPath("TestData")& "\"& sContent &".CATPart"
		If Fn_FSOUtil_FileOperations("fileexist",sFolderPath,"","")=True Then 
			sDestinationPath=Fn_Setup_GetAutomationFolderPath("tcic_tmp_Import") & "\"& sContent &".CATPart"
			If Fn_FSOUtil_FileOperations("copyfile",sFolderPath ,"",sDestinationPath)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to copy CAT Part from location [" & Fn_Setup_GetAutomationFolderPath("TestData")& " ] and paste under folder [ " & Cstr(sDestinationPath) & " ]","","","","","")				
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to copy CAT Part from [" & Fn_Setup_GetAutomationFolderPath("TestData")& " ] folder as CAT Part does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	
End Select


Function Fn_ExitTest()
	Set objShell = Nothing
	Set objClipboard = Nothing
	ExitTest
End Function

