 Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name											|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -|- - - - - - - - - - - - - - - -| - - - - - - - |- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_AWC_ReadyStatusSync									|	sandeep.navghane@sqs.com		|	26-Oct-2016	|	Function used to waits till Application comes to Ready state
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_AWC_ReadyStatusSync
'
'Function Description	 :	Function used to waits till AWC Application comes to Ready state
'
'Function Parameters	 :  1.iIterations: No. of times to be checked for Ready text						
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	AWC application should be displayed
'
'Function Usage		     :	Call Fn_AWC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-Oct-2016	    |	 1.0		|	  Prasenjeet P.	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_AWC_ReadyStatusSync(iIterations)
	Dim iCounter,iCount
	Dim bFlag
	Dim objAWCDefaultPage
	
	Fn_AWC_ReadyStatusSync=False
	bFlag=False
	Call Fn_Setup_ReporterFilter("DisableAll")
	
	Set objAWCDefaultPage=Browser("creationtime:=0" , "title:=.*Teamcenter.*").Page("title:=.*Teamcenter.*")
	objAWCDefaultPage.Sync
	
	Call Fn_Setup_ReporterFilter("EnableAll")
	For iCounter=1 To iIterations
		'Checking Existance of UserName link
		If Fn_WEB_UI_WebObject_Operations("Fn_AWC_ReadyStatusSync","Exist",objAWCDefaultPage.WebElement("class:=.*aw-state-userName.*","html tag:=A"),6,"","") Then
			bFlag=True
			wait 1
			Exit For
		End If
	Next
	IF bFlag=False Then
		Exit Function
	End IF
	For iCounter = 1 to iIterations
		bFlag=False
		If Fn_WEB_UI_WebObject_Operations("Fn_AWC_ReadyStatusSync","Exist",objAWCDefaultPage,6,"","") Then
			For iCount = 1 To 40
				If Fn_WEB_UI_WebObject_Operations("Fn_AWC_ReadyStatusSync","Exist",objAWCDefaultPage.WebElement("class:=aw-layout-progressBarCylon","html tag:=DIV"),2,"","") Then
					On Error Resume Next
					Call Fn_Setup_ReporterFilter("DisableAll")
					If Cint(objAWCDefaultPage.WebElement("class:=aw-layout-progressBarCylon","html tag:=DIV").GetROProperty("height"))=0 Then
						bFlag=True
						Exit For
					Else
						wait 1
					End If					
					Call Fn_Setup_ReporterFilter("EnableAll")
					On Error GoTo 0
				Else
					bFlag=False
					Exit For	
				End If	
			Next			
		Else
			Exit Function	
		End If
	Next
	IF bFlag=True Then
		Fn_AWC_ReadyStatusSync=True
	End IF
End Function
