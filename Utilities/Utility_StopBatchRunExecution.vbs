Call Fn_Utility_KillProcess("QTReport.exe:UFT.exe:QTP.exe:QTPAutomationAgent.exe:Javaw.exe:java.exe:Teamcenter.exe:EXCEL.exe:Acrord32.exe")

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
'Function Usage		     :  bReturn = Fn_Utility_KillProcess("Javaw.exe:java.exe:Teamcenter.exe:QTReport.exe:EXCEL.exe:Acrord32.exe:UFT.exe:QTP.exe:QTPAutomationAgent.exe")
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
	
	objWscriptShell.Popup  "Stoped batch run execution",5,"Stop Batch Run Execution"
	
	'Releasing object of wscript shell
	Set objWscriptShell = Nothing
End Function