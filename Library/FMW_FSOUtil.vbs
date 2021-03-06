Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name								|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_FSOUtil_FileOperations							|	vrushali.sahare@sqs.com		|	13-Jan-2015	|	Function Used to perform operations on local Files
'002. Fn_FSOUtil_FolderOperations						|	varada.satawalekar@sqs.com	|	13-Jan-2015 |	Function Used to perform folder related operations
'003. Fn_FSOUtil_ZipFileOperations						|	mandeep.deshwal@sqs.com		|	13-Jan-2015	|	Function used to perform operations on Zip files.
'004. Fn_FSOUtil_XMLFileOperations						|	ganesh.bhosale@sqs.com		|	13-Jan-2015	|	Function Used to perform operations on XML file
'005. Fn_FSOUtil_CreateExcelFromTextFile				|	kundan.kudale@sqs.com		|	05-Jun-2015	|	Function Used to save a text file as excel file.
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			:	Fn_FSOUtil_FileOperations
'
'Function Description	:	Function Used to perform operations on local file system text files.
'
'Function Parameters	:   1.sAction:Action name
'							2.sFilePath: File path or folder
'							3.sContent: File contents
'							4.sValue: New values
'
'Function Return Value	 : 	True or False / Date / File Size
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	File should exist in all cases except "createfile"
'
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("fileexist","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("modifytext","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","modified")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("verifytext","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","system","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("createfile","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("deletefile","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("getfilecreationdate","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("getlastmodifieddate","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("getfilesize","C:\VSEM_AUTOMATION\TestData\RunTimeTemp\CommonTestData.txt","","")
'Function Usage		     : 	bReturn = Fn_FSOUtil_FileOperations("deleteallfiles","C:\VSEM_AUTOMATION\TestData\RunTimeTemp","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  13-Jan-2015	    |	 1.0		|		Kundan Kudale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_FSOUtil_FileOperations(sAction,sFilePath,sContent,sValue)
	Err.Clear
	'Declaring variables
	Dim objFSO, objFile
	Dim sTextLine,objLastModifiedFile
	
	'Initially set function return value as False
	Fn_FSOUtil_FileOperations = False
	
	'Creating Object of File System
	Set objFSO = CreateObject("Scripting.FileSystemObject")

    Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to check specific file exist on specific location
		Case "fileexist"
			'If File exist then return True value else write failure log.
			If objFSO.FileExists(sFilePath) Then
				Fn_FSOUtil_FileOperations = True
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] as file does not exist at location [ " & Cstr(sFilePath) &" ]")
			End If	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to modify text in text file
		Case "modifytext"
			'Checking file existence
			If objFSO.FileExists(sFilePath) Then
				'If file exist then modify the text
				Set objFile = objFSO.OpenTextFile(sFilePath , 8  ,True)
				objFile.WriteLine sValue
				Fn_FSOUtil_FileOperations = True
				Set objFile = Nothing
			Else
				'If file does not exist then print failure log
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] as file does not exist at location [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify text in text file
		Case "verifytext"
			'Checking File exist or not
			If objFSO.FileExists(sFilePath) Then
				'If file exist then open the file for reading
				Set objFile = objFSO.OpenTextFile(sFilePath, 1, True )
				sTextLine = ""
				'Get all text file content in variable 
				Do While objFile.AtEndOfStream <> True
					sTextLine = sTextLine & objFile.ReadAll
				Loop
				'Check if expected value is present in text file.
				If Instr(1,sTextLine,sContent) Then
					Fn_FSOUtil_FileOperations = True
				End If
				Set objFile = Nothing
			Else
				'If file does not exist then print failure log
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] as file does not exist at location [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to create text file
		Case "createfile"
			'Check if file with same name already exist at given location.
			If not objFSO.FileExists(sFilePath) Then
				'If file is not already present then create a new file
				Set objFile = objFSO.CreateTextFile(sFilePath)
			Else
				'If file already present then delete existing file and then create a new file.
				objFSO.DeleteFile(sFilePath)
				Set objFile = objFSO.CreateTextFile(sFilePath)
			End If
			'Verify if created file exist at desired location
		   	If objFSO.FileExists(sFilePath) = True Then
				Fn_FSOUtil_FileOperations = sFilePath
			Else
				'Report failure if file creation was unsuccessful
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : failed to create file at location [ " & Cstr(sFilePath) &" ]")
			End If			
		 ' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to delete text file
		Case "deletefile"
			'Verify if file to be deleted is present at given location
			If objFSO.FileExists(sFilePath) = True Then
				'If file is present then delete the file
				objFSO.DeleteFile(sFilePath)
				'Verify file existence after deletion
				 If objFSO.FileExists(sFilePath) = False Then
					Fn_FSOUtil_FileOperations = True
				 Else
						'Report failure if file still exist after delete operation
						Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to delete file at location [ " & Cstr(sFilePath) &" ]")
				 End If	
			Else
				'Report failure if file to be deleted does not exist at given location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find file at location [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to retrieve date & time of file creation
		Case "getfilecreationdate"
			'Check file existence
			If objFSO.FileExists(sFilePath) Then
				'If file is present then get the date of creation of that file
				Set objFile = objFSO.GetFile(sFilePath)
				Fn_FSOUtil_FileOperations = objFile.DateCreated
				Set objFile = Nothing
			Else
				'Report failure of file not found at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find file at location [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to retrieve last modified date & time of file
		Case "getlastmodifieddate"
			'Check file existence
			'If objFSO.FileExists(sFilePath) Then
			
			If objFSO.FolderExists(sFilePath) Then
				'If file is present then get the last modified date and time
				Set objFile = objFSO.GetFile(sFilePath)
				Fn_FSOUtil_FileOperations = objFile.DateLastModified
				Set objFile = Nothing
			Else
				'Report failure of file not found at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find file at location [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to last modified file name in a folder
		Case "getlastmodifiedfilename"
			Fn_FSOUtil_FileOperations=""
			'Checking Folder existence
			Set objLastModifiedFile=Nothing
			If objFSO.FolderExists(sFilePath) Then
				For Each objFile In objFSO.GetFolder(sFilePath).Files
					If objLastModifiedFile Is Nothing Then 
						Set objLastModifiedFile = objFile
					 Else
						If objLastModifiedFile.DateLastModified < objFile.DateLastModified Then
						   Set objLastModifiedFile = objFile
						End If
					 End If
					Fn_FSOUtil_FileOperations = True
				Next
				If objLastModifiedFile Is Nothing Then
					Fn_FSOUtil_FileOperations=""
				Else
					Fn_FSOUtil_FileOperations=objLastModifiedFile.Name
				End If
			Else
				'Report failure of folder not found at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find folder [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to retrieve size of file in bytes
		Case "getfilesize"
			'Check file existence
			If objFSO.FileExists(sFilePath) Then
				'If file exist then get the size of file
				Set objFile = objFSO.GetFile(sFilePath)
				Fn_FSOUtil_FileOperations = objFile.Size
				Set objFile = Nothing
			Else
				'Report failure of file not found at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find file at location [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to delete all files in a folder
		Case "deleteallfiles"
			'Checking Folder existence
			If objFSO.FolderExists(sFilePath) Then
				'If folder exist then delete all the files present in that folder
				objFSO.DeleteFile(sFilePath & "\*.*"),True
				Fn_FSOUtil_FileOperations = True
			Else
				'Report failure of folder not found at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find folder [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to copy file
		Case "copyfile"
			'Checking Folder existence
			If objFSO.FileExists(sFilePath) Then
				objFSO.CopyFile sFilePath,sValue
				Fn_FSOUtil_FileOperations = True
			Else
				'Report failure of folder not found at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find folder [ " & Cstr(sFilePath) &" ]")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : No valid case was passed for function [Fn_FSOUtil_FileOperations]")
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release objects
	Set objFSO = Nothing
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - -- - - - - - - - - - - - - - - - - - -@Function Header Start 
'Function Name			:	Fn_FSOUtil_FolderOperationss
'
'Function Description	:	Function used to perform operations on file system folders
'
'Function Parameters	:   1.sAction: Action Name
'							2.sFolderPath:  Folder Path
'							3.sDestinationPath: Destination path in case of Copy, Move, Rename
'							4.sShareName: Desktop sharing Name for Share
'							5.sCompName: Name of computer/IP Address
'
'Function Return Value	: 	Folder names in string format \ True \ False
'
'Wrapper Function		: 	NA
'
'Function Pre-requisite	:	Folder on which operation is to be performed should exist in all cases except case "Create"
'
'Function Usage			:   bReturn = Fn_FSOUtil_FolderOperationss("Create", "C:\SQS", "", "","")
'							bReturn = Fn_FSOUtil_FolderOperationss("move","C:\SQS","D:\SQS","","")
'							bReturn = Fn_FSOUtil_FolderOperationss("exist","C:\SQS","","","")'
'                       	bReturn = Fn_FSOUtil_FolderOperationss("delete","C:\SQS","","","")
'							bReturn = Fn_FSOUtil_FolderOperationss("share","C:\SQS","","SQS","")
'							bReturn = Fn_FSOUtil_FolderOperationss("copy","C:\SQS","D:\ABC","","")
'							bReturn = Fn_FSOUtil_FolderOperationss("rename","C:\SQS","C:\ABC","","")
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Varada Satawalekar		 	| 13-Jan-2015	|	1.0			|	Kundan Kudale	 			| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header End

Public Function Fn_FSOUtil_FolderOperations(sAction, sFolderPath, sDestinationPath, sShareName,sCompName)
	'Declaring variables
	Dim objFSO,objFolder,objWMIService, objNewShare, objSubFolders, objCurrentSubFolder
	Dim iCounter
	Dim sLastModifiedFolder,sLastModifiedDate
	
	'Initially set function return value as false
	Fn_FSOUtil_FolderOperations = False
	
	'Create object of FileSystem
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Select case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to create a new folder
		Case "create"
			'Check if folder with same name already exist at desired location
			If objFSO.FolderExists(sFolderPath) Then
				'If folder already exist then delete that folder and create a new folder
				objFSO.DeleteFolder(sFolderPath)
				Set objFolder = objFSO.CreateFolder(sFolderPath)
			Else
				'Creating new folder with given name
				Set objFolder = objFSO.CreateFolder(sFolderPath)
			End If	
			
			'Verify if newly created folder exist
			If objFSO.FolderExists(sFolderPath) Then
				'If newly created folder exist then return true.
				Fn_FSOUtil_FolderOperations = True
			Else
				'If newly created folder does not exist then report failure
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to create folder at location [ " & Cstr(sFolderPath) &" ]")
			End If
			'Release objects from memory
			Set objFolder = Nothing
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to verify existence of folder at given location
		Case "exist"
			'Check if folder is present at desired location
			If objFSO.FolderExists(sFolderPath) Then
				'If folder is present then return true
				Fn_FSOUtil_FolderOperations = true
			Else
				'If folder not found at desired location then report failure
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ]")
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to delete a folder
		Case "delete"
			'Check if folder to be deleted is present at desired location
			If objFSO.FolderExists(sFolderPath) Then
				'If folder is found at desired location then delete the folder
				objFSO.DeleteFolder(sFolderPath)
				'Verify if deleted folder still exist
				If objFSO.FolderExists(sFolderPath) = False Then
					'If folder not found after deletion then return true
					Fn_FSOUtil_FolderOperations = true
				Else
					'If folder still exist after performing delete operation then report failure
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to delete folder at location [ " & Cstr(sFolderPath) &" ]")
				End If
			Else
				'If folder to be deleted is not present at desired location then report failure
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ] before performing delete operation")
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to share a given folder on network
		Case "share"
			'Declare constant variables
			Const FILE_SHARE = 0 
			Const MAXIMUM_CONNECTIONS = 25 
			
			'Check if folder to be shared is present at desired location
			If objFSO.FolderExists(sFolderPath) Then
				'If folder is present at desired location then share the folder
				Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & sCompName & "\root\cimv2")
				Set objNewShare = objWMIService.Get("Win32_Share")
				Fn_FSOUtil_FolderOperations = objNewShare.Create(sFolderPath, sShareName, FILE_SHARE, MAXIMUM_CONNECTIONS, sShareName)
			Else	
				'If folder to be shared is not present at desired location then report failure
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ] before performing sharing operation")
			End If

		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to copy a folder
		Case "copy"
			'Check if folder to be copied is present at desired location
			If objFSO.FolderExists(sFolderPath) Then
				'If folder is present then copy the folder at given path
				objFSO.CopyFolder sFolderPath,sDestinationPath,True
				Fn_FSOUtil_FolderOperations = True
			Else
				'If folder not present at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ] before performing copy operation")
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to move or rename a folder
		Case "move", "rename"
			'Check if folder to be copied is present at desired location
			If objFSO.FolderExists(sFolderPath) Then
				'Move the folder at given destination/Rename the folder
				objFSO.MoveFolder sFolderPath,sDestinationPath
				
				'Check if moved or renamed folder exist at destination folder
				If objFSO.FolderExists(sDestinationPath) Then
					Fn_FSOUtil_FolderOperations = True
				Else
					'If renamed or moved folder not present at destination path then report failure
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sDestinationPath) &" ] after performing [" & sAction & "] operation")
				End If
			Else
				'If folder not present at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ] before performing [" & sAction & "] operation")
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get file count of specific folder
		Case "getfilecount"
			'Check if folder to be copied is present at desired location
			If objFSO.FolderExists(sFolderPath) Then
				Set objFolder = objFSO.GetFolder(sFolderPath)
				Fn_FSOUtil_FolderOperations = objFolder.Files.Count	
			Else
				Fn_FSOUtil_FolderOperations=-1
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ] before performing [" & sAction & "] operation")
			End If	
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get name of all sub folders in a folder
		Case "subfolders"
			'Check parent folder existence
			If objFSO.FolderExists(sFolderPath) Then
				'If parent folder present then get all sub folders of given folder in collection format
				set objFolder = objFSO.GetFolder(sFolderPath)
				Set objSubFolders = objFolder.SubFolders
				
				'get name of all sub folders in one string separated by ~
				iCounter = 0
				For Each objCurrentSubFolder In objSubFolders
					If iCounter = 0 Then
						sSubFolderNames = objCurrentSubFolder.name
					Else
						sSubFolderNames = sSubFolderNames & "~" & objCurrentSubFolder.name
					End If
					iCounter = iCounter + 1
				Next
				Fn_FSOUtil_FolderOperations = sSubFolderNames
			Else
				'If parent folder not present at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ]")
			End If
			' release objects
			Set objFolder = Nothing
			Set objSubFolders = Nothing
			Set objCurrentSubFolder = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get name of all sub folders in a folder
		Case "getlastmodifiedsubfolder"
			'Check parent folder existence
			If objFSO.FolderExists(sFolderPath) Then
				'If parent folder present then get all sub folders of given folder in collection format
				set objFolder = objFSO.GetFolder(sFolderPath)
				Set objSubFolders = objFolder.SubFolders
				
				'get last modified folder name
				For Each objCurrentSubFolder In objSubFolders
					If objCurrentSubFolder.DateLastModified > sLastModifiedDate or isempty(sLastModifiedDate) Then
						sLastModifiedFolder = objCurrentSubFolder.Name
						sLastModifiedDate = objCurrentSubFolder.DateLastModified	
					End If
				Next
				Fn_FSOUtil_FolderOperations = sLastModifiedFolder
			Else
				'If parent folder not present at desired location
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of folder at location [ " & Cstr(sFolderPath) &" ]")
			End If
			' release objects
			Set objFolder = Nothing
			Set objSubFolders = Nothing
			Set objCurrentSubFolder = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : No valid case was passed for function [Fn_FSOUtil_FolderOperations]")
	
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_FolderOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release objects
	Set objFSO = Nothing
	
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			:	Fn_FSOUtil_ZipFileOperations
'
'Function Description	:	Function used to perform operations on Zip files.
'
'Function Parameters	:   1. sAction: Action Name
'							2. sFileLocation: File or folder path to unzip/zip
'							3. sExtractToLocation: Location to extract Zip file
'
'Function Return Value	: 	True or False
'
'Wrapper Function		: 	NA
'
'Function Pre-requisite	:	File or folder should be present.
'
'Function Usage			:   bReturn = Fn_FSOUtil_ZipFileOperations("Unzip","C:\SQS\AutomationXML.zip","C:\SQS")
'                       
'History			:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Mandeep Deshwal		 		| 13-Jan-2016	|	1.0			|	Kundan Kudale				| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header End
Public Function Fn_FSOUtil_ZipFileOperations(sAction,sFileLocation,sExtractToLocation)

 	'Declaring variables
	Dim objFSO,objShell,objZipItems

	'Initially set function return value as false
	Fn_FSOUtil_ZipFileOperations = False
	
	'Creating object of File System
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	'Creating shell object
	Set objShell = CreateObject("Shell.Application")

	Select Case Lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case the zip an unzipped folder or file
		Case "zip"
			'For Future Use	
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to unzip a file
		Case "unzip"
			'Checking existence of zip file
			If objFSO.FileExists(sFileLocation) Then
				'Checking existence of folder to unzip file
				If objFSO.FolderExists(sExtractToLocation) Then
					'Creating object of zipped items and copying items to destination path
					Set objZipItems = objShell.NameSpace(sFileLocation).Items
					objShell.NameSpace(sExtractToLocation).CopyHere(objZipItems)
					'Set function return value to True
					Fn_FSOUtil_ZipFileOperations = True
				Else
					'If destination folder does not exist then report failure
					Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_ZipFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of destination folder at location [ " & Cstr(sExtractToLocation) &" ]")
				End If
			Else
				'If zip file does not exist then report failure
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_ZipFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of zip file at location [ " & Cstr(sFileLocation) &" ]")
			End If
			'Releasing objects
			Set objZipItems=Nothing
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_ZipFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : No valid case was passed for function [Fn_FSOUtil_ZipFileOperations]")
		
	End Select
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_ZipFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Release objects from memory
	Set objShell=Nothing
	Set objFSO=Nothing	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - @Function Header Start
'Function Name			:	Fn_FSOUtil_XMLFileOperations
'
'Function Description	:	Function Used to perform operations on XML file
'
'Function Parameters	:   1.sAction: Action Name
'							2.sXMLFilePath: XML file path \  XML File Name for 'getobject' Action 
'							3.sNodeName: Variable name for which value needs to be fetched or set 
'							4.sNodeValue: Node Value for the sNodeName 
'
'Function Return Value	: 	True \ False \ Value of node
'
'Wrapper Function		: 	NA
'
'Function Pre-requisite	:	XML File should exist.
'
'Function Usage			:   bReturn = Fn_FSOUtil_XMLFileOperations("getvalue","C:\Automation_Mainline\AutomationXML\MenuXML\RAC_Menu.xml","FileExit","")
'Function Usage			:   bReturn = Fn_FSOUtil_XMLFileOperations("setvalue","C:\Automation_Mainline\AutomationXML\SetupXML\EnvironmentVariables.xml","BrowserName","IE")
'Function Usage			:   bReturn = Fn_FSOUtil_XMLFileOperations("getobject","EnvironmentVariables","TcDefaultApplet","")
'Function Usage			:   bReturn = Fn_FSOUtil_XMLFileOperations("getallnodevalues","EnvironmentVariables","","")
'                       
'History				:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Ganesh Bhosale			 	| 13-Jan-2016	|	1.0			|	Kundan Kudale		 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header End
Public Function Fn_FSOUtil_XMLFileOperations(sAction,ByVal sXMLFilePath,sNodeName,sNodeValue)
	Err.Clear
	'Declaring variables
	Dim objXMLDOM, objCurrentNameNode, objCurrentValueNode
	Dim iTotalNumberOfVariables, iCounter
	Dim bFlag
	
	'Initially set function return value as False
	Fn_FSOUtil_XMLFileOperations = False
	bFlag = False
	
	'Create XMLDOM object
	Set objXMLDOM = CreateObject("Microsoft.XMLDOM")												
	objXMLDOM.async="false"
	
	' get File Path
	sXMLFilePath=Fn_Setup_GetAutomationXMLPath(sXMLFilePath)
	
	'Check if XML file exist
	IF Fn_FSOUtil_FileOperations("fileexist",sXMLFilePath,"","") = False Then
		'If XML file on which operations are to be performed is not present at desired location then report failure
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to verify existence of XML file at location [ " & Cstr(sXMLFilePath) &" ]")
		Set objXMLDOM = Nothing
		Exit Function
	Else
		'If XML file exist then load the XML file
		objXMLDOM.Load(sXMLFilePath)
	End IF
	
	'Report failure if the XML file contains any errors
	If (objXMLDOM.parseError.errorCode <> 0) Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : XML file at location [ " & Cstr(sXMLFilePath) &" ] contains errors.")
		Set objXMLDOM = Nothing
		Exit Function
	End If
	
	Select Case Lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get value of a node in XML file
		Case "getvalue"
			Set objCurrentNameNode=objXMLDOM.selectNodes("//Environment/Variable[@Name='" & sNodeName & "']")
			If objCurrentNameNode.length=1 Then
				Fn_FSOUtil_XMLFileOperations = Trim(objCurrentNameNode.Item(0).Text)
				bFlag = True
			End If

			If bFlag = False Then	
				If objXMLDOM.getElementsByTagName("Value").length=0 Then
					bFlag = False
				Else				
					'Get the total number of variable tags present in XML file
					iTotalNumberOfVariables = objXMLDOM.getElementsByTagName("Variable").length
					
					'Iterate a loop for each variable node present in XMl file
					For iCounter = 0 to (iTotalNumberOfVariables - 1)
						'Get name tag object for current variable
						Set objCurrentNameNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Name")
						
						'Compare the current name node value with expected node name
						If Trim(objCurrentNameNode.Text) = Trim(sNodeName) Then
							'If current name node value matches expected then get the value tag object
							Set objCurrentValueNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Value")
							'Return the current value node's value
							Fn_FSOUtil_XMLFileOperations = Trim(objCurrentValueNode.Text)
							bFlag = True
							Exit For
						End If
					Next
				End If
			End If
			'If expected node found then report failure 
			If bFlag = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find node [ " & Cstr(sNodeName) &" ] in XML at path [" & Cstr(sXMLFilePath) & "].")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get object hierarchy from XML file
		Case "getobject"
			bFlag = ""
			
			Set objCurrentNameNode=objXMLDOM.selectNodes("//Environment/Variable[@Name='" & sNodeName & "']")
			If objCurrentNameNode.length=1 Then
				bFlag = Trim(objCurrentNameNode.Item(0).Text)
			End If
			
			If bFlag = "" Or bFlag = False Then 
				'Get the total number of variable tags present in XML file
				iTotalNumberOfVariables = objXMLDOM.getElementsByTagName("Variable").length
				
				'Iterate a loop for each variable node present in XMl file
				For iCounter = 0 to (iTotalNumberOfVariables - 1)
					'Get name tag object for current variable
					Set objCurrentNameNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Name")
					
					'Compare the current name node value with expected node name
					If Trim(objCurrentNameNode.Text) = Trim(sNodeName) Then
						'If current name node value matches expected then get the value tag object
						Set objCurrentValueNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Value")
						'Return the Object
						bFlag= objCurrentValueNode.Text
						Exit For
					End If
				Next
			End IF
			'If expected node not found then report failure or return object
			If bFlag <> "" AND bFlag <> False Then 
				Set Fn_FSOUtil_XMLFileOperations = eval(bFlag)
			Else
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find node [ " & Cstr(sNodeName) &" ] in XML at path [" & Cstr(sXMLFilePath) & "].")
			End If
			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set value of a node in XML file
		Case "setvalue"
			Set objCurrentNameNode=objXMLDOM.selectNodes("//Environment/Variable[@Name='" & sNodeName & "']")
			If objCurrentNameNode.length=1 Then
				objCurrentNameNode.Item(0).Text= sNodeValue
				objXMLDOM.Save(sXMLFilePath)
				bFlag =True
			End If
			
			If bFlag = False Then
				'Get the total number of variable tags present in XML file
				iTotalNumberOfVariables = objXMLDOM.getElementsByTagName("Variable").length
				
				'Iterate a loop for each variable node present in XMl file
				For iCounter = 0 to (iTotalNumberOfVariables - 1)
					'Get name tag object for current variable
					Set objCurrentNameNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Name")
					
					'Compare the current name node value with expected node name
					If Trim(objCurrentNameNode.Text) = Trim(sNodeName) Then
					'If current name node value matches expected then get the value tag object
						Set objCurrentValueNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Value")
						'Set the current value node's value
						objCurrentValueNode.Text = sNodeValue											
						'Save XML file
						objXMLDOM.Save(sXMLFilePath)
						bFlag = True
						Exit For
					End If
				Next
			End IF
			
			'If expected node found then report failure 
			If bFlag = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to find node [ " & Cstr(sNodeName) &" ] in XML at path [" & Cstr(sXMLFilePath) & "].")
			Else
				Fn_FSOUtil_XMLFileOperations = True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get all node values from xml
		Case "getallnodevalues"
			If objXMLDOM.getElementsByTagName("Value").length=0 Then
				iTotalNumberOfVariables = objXMLDOM.getElementsByTagName("Variable").length
				sNodeValue=""
				'Iterate a loop for each variable node present in XMl file
				For iCounter = 0 to (iTotalNumberOfVariables - 1)
					'Get name tag object for current variable
					Set objCurrentValueNode = objXMLDOM.SelectNodes("//Environment/Variable")				
					'Retrive node value
					If iCounter = 0 Then
						sNodeValue=Trim(objCurrentValueNode.item(iCounter).Text)
					Else
						sNodeValue=sNodeValue & "~" & Trim(objCurrentValueNode.item(iCounter).Text)
					End If
				Next
			Else
				'Get the total number of variable tags present in XML file
				iTotalNumberOfVariables = objXMLDOM.getElementsByTagName("Variable").length
				sNodeValue=""
				'Iterate a loop for each variable node present in XMl file
				For iCounter = 0 to (iTotalNumberOfVariables - 1)
					'Get name tag object for current variable
					Set objCurrentValueNode = objXMLDOM.SelectSingleNode("/Environment/Variable[" & iCounter &"]/Value")				
					'Retrive node value
					If iCounter = 0 Then
						sNodeValue=Trim(objCurrentValueNode.Text)
					Else
						sNodeValue=sNodeValue & "~" & Trim(objCurrentValueNode.Text)
					End If
				Next
			End If
			If sNodeValue<>"" Then
				Fn_FSOUtil_XMLFileOperations = sNodeValue
			Else
				Fn_FSOUtil_XMLFileOperations=False
			End If

			If Fn_FSOUtil_XMLFileOperations = False Then
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : Failed to read all values from XML at path [" & Cstr(sXMLFilePath) & "].")
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] : No valid case was passed for function [Fn_FSOUtil_XMLFileOperations]")
		
	End Select
	
	'Release all Objects
	Set objCurrentNameNode = Nothing 
	Set objXMLDOM = Nothing
	Set objCurrentValueNode = Nothing 
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_FSOUtil_XMLFileOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
End Function


'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -- - - @Function Header Start
'Function Name			:	Fn_FSOUtil_CreateExcelFromTextFile
'
'Function Description	:	Function Used to create an Excel file from text file
'
'Function Parameters	:   1.sTextFilePath: Source text file from which data is to be referred
'							2.sExcelFilePath: Target xls file path where data is to be saved
'							3.sDelimiterValue: Delimiter value using which text data is to be placed on excel cells
'
'Function Return Value	: 	True \ False
'
'Wrapper Function		: 	NA
'
'Function Pre-requisite	:	Text file should exist
'
'Function Usage			:   bReturn = Fn_FSOUtil_CreateExcelFromTextFile("C:\abc.txt", "C:\abc.xls", vbTab)
'                       
'History				:
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Developer Name				|	Date		|	Rev. No.	|	Reviewer					|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'	Kundan Kudale			 	| 05-Jun-2017	|	1.0			|	Sabdeep Navghane	 		| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header End

Public Function Fn_FSOUtil_CreateExcelFromTextFile(sTextFilePath, sExcelFilePath, sDelimiterValue)

	'Variable declaration
	Dim objExcel
	
	On Error Resume Next
	
	'Create an Excel application object
	Set objExcel = CreateObject("Excel.Application")
	
	'Open the text file using the delimiter specified
	objExcel.Workbooks.Open sTextFilePath,,,6,,,,,sDelimiterValue
	
	'Save a copy of the text file opened in excel applciation
	objExcel.ActiveWorkbook.SaveCopyAs(sExcelFilePath)
	
	'Close the excel workbook
	objExcel.ActiveWorkbook.Close
	
	'Close excel application
	objExcel.Quit
	Set objExcel = Nothing
	
	'Return function execution result
	If Err.Number <> 0 Then
		Fn_FSOUtil_CreateExcelFromTextFile = False
	Else
		Fn_FSOUtil_CreateExcelFromTextFile = True
	End If
	
End Function