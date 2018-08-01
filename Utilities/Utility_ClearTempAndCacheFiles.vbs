'Option Explicit
Call Fn_Utility_KillProcess("Javaw.exe:java.exe:Teamcenter.exe:QTReport.exe:EXCEL.exe:Acrord32.exe")
Call Fn_Utility_DeleteCacheFiles()
Call Fn_Utility_DeleteTempFolderAndFiles()	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name									|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_Utility_DeleteTempFolderAndFiles				|	Sandeep.Navghane@sqs.com	|	16-Mar-2015	|	Function used to delete temp folders and files
'002. 	Fn_Utility_DeleteCacheFiles						|	Sandeep.Navghane@sqs.com	|	16-Mar-2015	|	Function used to delete cache files
'003. 	Fn_Utility_KillProcess							|	Sandeep.Navghane@sqs.com	|	16-Mar-2015	|	Function used to terminate process
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Utility_DeleteTempFolderAndFiles
'
'Function Description	 :	Function used to delete temp folders and files
'
'Function Parameters	 :  NA
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_Utility_DeleteTempFolderAndFiles()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Utility_DeleteTempFolderAndFiles()
	On Error Resume Next
	'Declaring variables
	Dim objFolder, objSubFolder,objFSO,objFiles,objWscriptShell,objWscriptNetwork
	Dim sCurrentLoginUser
	DIm iCounter
	Dim fileIdx
	
	'Deleting local temp objFiles
	Set objFSO = CreateObject("Scripting.FilesystemObject")
	Set objWscriptShell = CreateObject("Wscript.Shell")
	Set objFolder = objFSO.GetFolder("C:\temp")
	Set objSubFolder = objFolder.SubFolders
	For Each iCounter in objSubFolder
		Err.Clear
		objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True
	Next
	Set objFiles = objFolder.Files
	For each fileIdx In objFiles    
		fileIdx.Delete true
	Next		
	Set objFiles = Nothing
	Set objFSO = Nothing
	Set objFolder = Nothing 
	Set objSubFolder = Nothing
	
	Call Fn_Utility_DeleteCacheFiles()
	
	'Deleting user temp Files
	Set objFSO = CreateObject("Scripting.FilesystemObject")
	Set objWscriptShell = CreateObject("Wscript.Shell")
	Set objWscriptNetwork = Wscript.CreateObject("WScript.Network")
	Set objFolder = objFSO.GetFolder("C:\Users\" & objWscriptNetwork.UserName & "\AppData\Local\Temp")
	
	Set objSubFolder = objFolder.SubFolders
	For Each iCounter in objSubFolder
		objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True
	Next
	
	Set objFiles = objFolder.Files
	For Each iCounter In objFiles    
		objFSO.DeleteFile iCounter,True
	Next	
	Err.Clear
	
	'Releasing all objects
	Set objWscriptNetwork =Nothing
	Set objSubFolder = Nothing
	Set objFolder = Nothing 
	Set objFSO = Nothing
	
	Call Fn_Utility_DeleteCacheFiles()
	
	objWscriptShell.Popup  "Deleted All Cache Files & Folders | Deleted All Temp Files & Folders",5,"Clear Temp and Cache Files\Folders"
	Set objWscriptShell=Nothing
End Function		
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Utility_DeleteCacheFiles
'
'Function Description	 :	Function used to delete cache files
'
'Function Parameters	 :  NA
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_Utility_DeleteCacheFiles()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_Utility_DeleteCacheFiles()
		On Error Resume Next
		'Variable Declaration
		Dim objFolder,objSubFolder,objWscriptShell,objFSO,objWscriptNetwork
		Dim sFolderName
		Dim aFolderName
		Dim iCounter
		
		'Constant Variable Declaration
		Const DeleteReadOnly = True
		'Creating Object of Shell
		Set objWscriptShell = CreateObject("Wscript.Shell")
		'Creating Object of Filesystem
		Set objFSO = CreateObject("Scripting.FilesystemObject") 
		'Creating Object of Network
		Set objWscriptNetwork = Wscript.CreateObject("WScript.Network")		
		Set objFolder = objFSO.GetFolder("C:\Documents and Settings\" & objWscriptNetwork.UserName)	
		
		sFolderName = "FCCCache"
		'Deleting FCCCache folders
		Set objSubFolder = objFolder.SubFolders
		For Each iCounter in objSubFolder
			aFolderName = split(iCounter.Name, "_",-1,1)
			if sFolderName = aFolderName(0) Then
			  objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True 		
			End if 
		Next
		'Deleting Teamcenter folders
		objFSO.DeleteFolder objFolder & "\" & "Teamcenter",True	
		'Deleting Siemens folders	
		objFSO.DeleteFolder objFolder & "\" & "Siemens",True 
		objFSO.DeleteFolder objFolder & "\" & ".TcIC",True 	
		objFSO.DeleteFolder objFolder & "\" & ".swt",True 
		sFolderName = ".Administrator"
		
		For Each iCounter in objSubFolder
			aFolderName = split(iCounter.Name, "_",-1,1)
			If sFolderName = aFolderName(0) then
				'Deleting .Administrator folders
				objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True		
			End if 
		Next
		
		For Each iCounter in objSubFolder
			aFolderName = split(iCounter.Name, "_",-1,1)
			'Check the existing of .Administrator and Teamcenter folders	
			If iCounter.Name = "Teamcenter" OR aFolderName(0)= ".Administrator" then	
				Call Fn_Utility_DeleteCacheFiles()
			End if
		Next			
		'Deleting Fcc Files
		objFSO.DeleteFile (objFolder & "\" & "fcc.*"),DeleteReadOnly
		objFSO.DeleteFile (objFolder & "\" & "TCPLM-JAVA.txt"),DeleteReadOnly
		
		Set objSubFolder = objFolder.SubFolders
		For Each iCounter in objSubFolder
			if Instr(1,iCounter.Name,"_lock_") Then
			  objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True 		
			End if 
		Next
		
		'Releasing all objects
		Set objWscriptNetwork = Nothing
		Set objWscriptShell = Nothing
		Set objSubFolder = Nothing
		Set objFolder = Nothing 
		Set objFSO = Nothing
End Function	
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Utility_KillProcess
'
'Function Description	 :	Function used to terminate process
'
'Function Parameters	 :  1.sProcessToKill	:  Process names to terminate
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_Utility_KillProcess("Javaw.exe:java.exe:Teamcenter.exe:QTReport.exe:EXCEL.exe:Acrord32.exe")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Utility_KillProcess(sProcessToKill)                                
	'Declaring variables
	Dim iCounter
	Dim sComputer
	Dim aProcessName
	Dim objWMIService,objProcess,objProcessCollection,objWscriptShell
	
	'Creating object of wscript shell
	Set objWscriptShell = CreateObject("Wscript.Shell")
	
	aProcessName = split(sProcessToKill,":",-1,1)
	sComputer = "." 
	Set objWMIService = GetObject("winmgmts:"& "{impersonationLevel=impersonate}!\\"& sComputer & "\root\cimv2") 
	'Terminating process
	For iCounter = 0 to ubound(aProcessName)
		Set objProcessCollection = objWMIService.ExecQuery("Select * from Win32_Process Where Name ='" & aProcessName(iCounter) & "'")
		For Each objProcess in objProcessCollection 
			objProcess.Terminate() 
		Next 
	Next
	'Releasing object of wscript shell
	Set objWscriptShell = Nothing
End Function