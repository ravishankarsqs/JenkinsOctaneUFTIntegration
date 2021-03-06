Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name											|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -|- - - - - - - - - - - - - - - -| - - - - - - - |- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_CATIA_ReadyStatusSync								|	sandeep.navghane@sqs.com		|	26-Oct-2016	|	Function used to waits till Application comes to Ready state
'002. 	Fn_CATIA_SetVisible										|	sandeep.navghane@sqs.com		|	26-Oct-2016	|	Function used to make CATIA Application visible\actiavte
'003. 	Fn_CATIA_WindowOperations								|	sandeep.navghane@sqs.com		|	26-Oct-2016	|	Function use to perform operations on window
'004. 	Fn_CATIA_TSMReadyStatusSync								|	sandeep.navghane@sqs.com		|	26-Oct-2016	|	Function used to waits till Teamcenter Save Manager Application comes to Ready state
'005. 	Fn_CATIA_GetTreeNodeObject								|	sandeep.navghane@sqs.com		|	26-Oct-2016	|	Function used to retrive catia tree node path\object by accessing CATIA API
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CATIA_ReadyStatusSync
'
'Function Description	 :	Function used to waits till CATIA Application comes to Ready state
'
'Function Parameters	 :  1.iIterations: No. of times to be checked for Ready text						
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	CATIA application should be displayed
'
'Function Usage		     :	Call Fn_AWC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-Oct-2016	    |	 1.0		|	  Prasenjeet P.	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Public Function Fn_CATIA_ReadyStatusSync(iIterations)
' 	'Declaring variables
'	Dim iCounter,iCount
'	Dim bFlag,bReturn
'	Dim sCurrentStatus
'	Dim objCATIA
'
'	bFlag =  false
'
'	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	'checking existance of CATIA default window
'	If Window("regexpwndtitle:=CATIA.*").Exist(2) Then
'		'checking current state of mouse
'        For iCounter =1 To iIterations
'			For iCount=1 To 25
'				bReturn=Fn_CommonUtil_GetCursorState()
'				If  bReturn="65543" OR bReturn="65561" Then  '//  Poiters for wait 
'					'wait GBL_MICRO_TIMEOUT
'					Wait 0,500
'				Else
'					'wait GBL_MICRO_TIMEOUT
'					Wait 0,500
'					Exit For
'				End If
'			Next
'		Next
'
'		'Checking status of CATIA application
'		For iCounter = 1 to iIterations
'			For iCount=0 to 25						
'					Set objCATIA=GetObject("","CATIA.Application")
'					sCurrentStatus=objCATIA.Statusbar
'
'					Select Case Lcase(sCurrentStatus)
'						Case "saving component in"
'							'wait GBL_MICRO_TIMEOUT
'							Wait 0,500
'						Case "checks-out in"
'							'wait GBL_MICRO_TIMEOUT
'							Wait 0,500
'						Case "loading merge"
'							'wait GBL_MICRO_TIMEOUT
'							Wait 0,500
'						Case "loading in catia"
'							'wait GBL_MICRO_TIMEOUT
'							Wait 0,500
'						Case "in teamcenter..."
'							'wait GBL_MICRO_TIMEOUT
'							Wait 0,500
'						Case Else
'							'wait GBL_MICRO_TIMEOUT
'							Wait 0,500
'							bFlag=True
'							Exit For
'					End Select
'
'					Set objCATIA=Nothing
'					'Validating any popup dialog appears
'					If Dialog("nativeclass:=#32770").Exist(1) Then
'						bFlag=True
'						'wait GBL_MICRO_TIMEOUT
'						Wait 0,500
'						Exit For
'					ElseIf Window("regexpwndtitle:=CATIA.*").Dialog("nativeclass:=#32770").Exist(1) Then
'						bFlag=True
'						'wait GBL_MICRO_TIMEOUT
'						Wait 0,500
'						Exit For
'					End IF
'			Next
'		Next
'	End If
'	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'
'	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	If bFlag = FALSE Then
'		Fn_CATIA_ReadyStatusSync = FALSE
'	Else
'		Fn_CATIA_ReadyStatusSync = TRUE
'	End If
'	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'End Function

Public Function Fn_CATIA_ReadyStatusSync(iIterations)
 	'Declaring variables
	Dim iCounter,iCount
	Dim bFlag,bReturn
	Dim sCurrentStatus
	Dim objCATIA

	bFlag =  false

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'checking existance of CATIA default window
	If Window("regexpwndtitle:=CATIA.*").Exist(0) Then
		'checking current state of mouse
        For iCounter =1 To iIterations
			For iCount=1 To 125
				bReturn=Fn_CommonUtil_GetCursorState()
				If  bReturn="65543" OR bReturn="65561" Then  '//  Poiters for wait 
					Wait 0,50
				Else
					Wait 0,50
					bFlag =  True
					Exit For
				End If
			Next
			If bFlag Then Exit for
		Next
		
		bFlag=False
		'Checking status of CATIA application
		For iCounter = 1 to iIterations
			For iCount=0 to 125		
				'Validating any popup dialog appears
				If Dialog("nativeclass:=#32770").Exist(0) Then
					bFlag=True
					Exit For
				ElseIf Window("regexpwndtitle:=CATIA.*").Dialog("nativeclass:=#32770").Exist(0) Then
					bFlag=True
					Exit For
				End IF
				
				Set objCATIA=GetObject("","CATIA.Application")
				sCurrentStatus=objCATIA.Statusbar

				Select Case Lcase(sCurrentStatus)
					Case "saving component in"
						Wait 0,50
					Case "checks-out in"
						Wait 0,50
					Case "loading merge"
						Wait 0,50
					Case "loading in catia"
						Wait 0,50
					Case "in teamcenter..."
						Wait 0,50
					Case Else
						Wait 0,50
						bFlag=True
						Exit For
				End Select
				Set objCATIA=Nothing
			Next
			If bFlag Then Exit for
		Next
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	If bFlag = FALSE Then
		Fn_CATIA_ReadyStatusSync = FALSE
	Else
		Fn_CATIA_ReadyStatusSync = TRUE
	End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CATIA_SetVisible
'
'Function Description	 :	Function used to make CATIA Application visible\actiavte
'
'Function Parameters	 :  NA						
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	CATIA application should be available
'
'Function Usage		     :	Call Fn_CATIA_SetVisible()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-Oct-2016	    |	 1.0		|	  Prasenjeet P.	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_CATIA_SetVisible()
   Dim iCounter
   If Window("regexpwndtitle:=CATIA.*").Exist(2)=False Then
		Exit Function
   End IF
   For iCounter=0 to 3
		If Window("regexpwndtitle:=CATIA.*").GetROProperty("visible") Then
		   Exit For
		End If
		Window("regexpwndtitle:=CATIA.*").highlight
		wait 2
		Window("regexpwndtitle:=CATIA.*").RefreshObject
   Next
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CATIA_WindowOperations
'
'Function Description	 :	Function use to perform operations on window
'
'Function Parameters	 :  1. sAction : Action name
'							2.sWindowTitle : Window title						
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Should be logged in to CATIA application
'
'Function Usage		     :	Call Fn_CATIA_WindowOperations("Maximize","DSV5A1.CATProduct")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-Oct-2016	    |	 1.0		|	  Prasenjeet P.	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_CATIA_WindowOperations(sAction,sWindowTitle)
	'declaring variables
	Dim objCATIA,objWindows,objActiveWindow
	Dim bFlag
	Dim iCounter

	Fn_CATIA_WindowOperations=False

	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -      
		Case "ActivateExt"
			Set objCATIA=GetObject("","CATIA.Application")
			Set objWindows = objCATIA.Windows
			Set objActiveWindow=objWindows.Item(sWindowTitle)
			wait 1
			'creating object of active window
			objActiveWindow.Activate
			Set objActiveWindow=Nothing  	
			Set objWindows = Nothing
			Set objCATIA=Nothing
			If Err.Number<0 Then
				Fn_CATIA_WindowOperations=False
			Else
				Fn_CATIA_WindowOperations=True
			End If
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -      
	 	Case "Maximize","Activate"
			'Creating object of CATIAapplication
			Set objCATIA=GetObject("","CATIA.Application")
			'Setting title
			If sWindowTitle<>"" Then
				'Creating object of Windows
				Set objWindows = objCATIA.Windows
				bFlag=False
				For iCounter=1 to objWindows.Count
					If Trim(sWindowTitle)=Trim(objWindows.Item(iCounter).Name) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=False Then
					Set objWindows =Nothing
					Set objCATIA=Nothing
					Exit Function
				End If
				'Selecting required window to perform operations
				Set objActiveWindow=objWindows.Item(iCounter)
				wait 1
				'creating object of active window
				objActiveWindow.Activate
				If sAction="Maximize" Then
					objActiveWindow.WindowState = 0
				End If
				
				Fn_CATIA_WindowOperations=True

				'Releasing objects
				Set objActiveWindow=Nothing
				Set objWindows =Nothing
				Set objCATIA=Nothing
			End If
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CATIA_TSMReadyStatusSync
'
'Function Description	 :	Function used to waits till Teamcenter Save Manager Application comes to Ready state
'
'Function Parameters	 :  1.iIterations: No. of times to be checked for Ready text						
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcenter Save Manager application should be displayed
'
'Function Usage		     :	Call Fn_CATIA_TSMReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-Oct-2016	    |	 1.0		|	  Prasenjeet P.	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CATIA_TSMReadyStatusSync(iIterations)
 	'Declaring variables
	Dim iCount
	
	Fn_CATIA_TSMReadyStatusSync=False
	
	If iIterations="" Then
		iIterations=20
	End If

	'Checking status bar
	For iCount = 0 to iIterations
		If  JavaWindow("title:=Teamcenter Save Manager.*").JavaStaticText("path:=.*StatusBar.*").Exist(1) Then
			wait 1
		Else
			Exit For
		End If
	Next

	Fn_CATIA_TSMReadyStatusSync=True
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CATIA_GetTreeNodeObject
'
'Function Description	 :	Function used to retrive catia tree node path\object by accessing CATIA API
'
'Function Parameters	 :  1.sTreeType 		: Catia tree type (Product,Part,Hybrid)					
'							2.sWindowTitle 		: Catia current window title
'							3.sTreeNodePath		: Complete node path
'							4.sTreeNodeTypes	: Each node type ( Optional )
'							5.sDelimiter 		: Tree node delimiter
'							6.sInstanceHandler 	: Tree node Instance Handler
'
'Function Return Value	 : 	Nothing or Tree node object
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Catia tree node should be available
'							
'Function Usage		     :	sNodePath="FNA1784891_1~FNA0273842_1.1~FNA0273840_1.1~FNA0273840_1~External Reference~Surface.1"
'							sNodeTypes="Product~Product~Product~Part~ExternalReference~Surface"
'							bReturn=Fn_CATIA_GetTreeNodeObject("Hybrid","",sNodePath,sNodeTypes,"","")
'
'							bReturn=Fn_CATIA_GetTreeNodeObject("Product","","FNA1784891_1~FNA0273842_1.1~FNA0273840_1.1","","","")
'
'							bReturn=Fn_CATIA_GetTreeNodeObject("Part","","FNA1784891_1~FNA0273842_1.1~FNA0273840_1.1","","","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  11-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_CATIA_GetTreeNodeObject(sTreeType,sWindowTitle,sTreeNodePath,sTreeNodeTypes,sDelimiter,sInstanceHandler)
	'Declaring variables
	Dim objDocument,objSelection,objWindows
	Dim objCATIA,objPart,objBodies,objPublish,objProducts,objSheets,objViews
	Dim bFlag,bFound
	Dim iCounter,iCount,iNextStartCounter,iCount1,iCounter1
	Dim aNode,aNodeType,aNodePath,aPartName
	Dim sLastNode,sCurrentNode

	'Set initial value to nothing
	Set Fn_CATIA_GetTreeNodeObject=Nothing
	'Creating object of CATIAapplication
	Set objCATIA=GetObject("","CATIA.Application")
	'Setting title
	If sWindowTitle<>"" Then
		'Creating object of Windows
		Set objWindows = objCATIA.Windows
		bFlag=False
		For iCounter=1 to objWindows.Count
			If lcase(Trim(sWindowTitle))=lcase(Trim(objWindows.Item(iCounter).Name)) Then
				bFlag=True
				Exit For
			End If
		Next
		If bFlag=False Then
			Set objWindows =Nothing
			Set objCATIA=Nothing
			Exit Function
		End If
		'Selecting required window to perform operations
		objWindows.Item(iCounter).Activate 
		wait 1
		Set objWindows =Nothing
	End If
	
	Select Case sTreeType
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get product tree node object
		Case "Product"
			'creating object of Active document of CATIA application
			Set objDocument = objCATIA.ActiveDocument

			aNode=Split(sTreeNodePath,"~")
			'Mapping Actual value with expected values
			If aNode(0)<>Trim(objDocument.Product.Name) Then
				Set objCATIA=Nothing
				Exit Function
			End If

			If ubound(aNode)=0 Then
				Set Fn_CATIA_GetTreeNodeObject=objDocument.Product
				Set objCATIA=Nothing
				Exit Function
			End If

			Set objProducts= objDocument.Product.Products

			'Mapping Actual value with expected values

			For iCounter=1 to Ubound(aNode)
				bFound=False
				For iCount=1 to objProducts.Count
					If Trim(objProducts.Item(iCount).Name)=aNode(iCounter) Then
						If Ubound(aNode)=iCounter Then
							Set objProducts=objProducts.Item(iCount)
						Else
							Set objProducts=objProducts.Item(iCount).Products
						End If
						bFound=True
						Exit For
					End If
				Next
				If bFound=False Then
					Exit For
				End If
			Next
			'Function returns of object
			If bFound=True Then
				Set Fn_CATIA_GetTreeNodeObject=objProducts
			End If
			'Releasing object
			Set objProducts=Nothing
			Set objDocument =Nothing
			Set objCATIA=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get part tree node object
		Case "Part"
			'creating object of Active document of CATIA application
			Set objDocument = objCATIA.ActiveDocument
			
			aNode=Split(sTreeNodePath,"~")
			aNodeType=Split(sTreeNodeTypes,"~")

'			Mapping Actual value with expected values
			If aNode(0)<>Trim(objDocument.Part.Name) Then
				Set objCATIA=Nothing
				Exit Function
			End If

			If ubound(aNode)=0 Then
				Set Fn_CATIA_GetTreeNodeObject=objDocument.Part
				Set objCATIA=Nothing
				Exit Function
			End If
			'Creating object of part
			Set objPart=objDocument.Part
			
			sLastNode= aNode(Ubound(aNode))
			
			If sLastNode="xy plane" or sLastNode="zx plane" or sLastNode="yz plane" Then
				Select Case sLastNode
					Case "xy plane"
						Set objPart=objPart.OriginElements.PlaneXY
					Case "zx plane"
						Set objPart=objPart.OriginElements.PlaneZX
					Case "yz plane"
						Set objPart=objPart.OriginElements.PlaneYZ
				End Select
				Set Fn_CATIA_GetTreeNodeObject=objPart
				Set objPart=Nothing
				Set objCATIA=Nothing
				Exit Function
			End If
								
			'Mapping Actual value with expected values
			For iCounter=1 to Ubound(aNode)
				bFound=False
				Select Case Lcase(aNodeType(iCounter))
					Case "bodies"
						Set objPart=objPart.Bodies
					Case "shapes"
						Set objPart=objBodies
						Set objPart=objPart.Shapes
					Case "sketches"
						Set objPart=objBodies
						Set objPart=objPart.Sketches
					Case "geometricelements"
						Set objPart=objPart.GeometricElements
					Case "constraints"
						Set objPart=objPart.Constraints
					Case "relations"
						Set objPart=objPart.Relations
					Case "externalreferences"
						Set objPart= objPart.HybridBodies
					Case "surface"
						Set objPart = objPart.HybridShapes					
				End Select
				

				For iCount=1 to objPart.Count
					If Trim(objPart.Item(iCount).Name)=aNode(iCounter) Then

						If Lcase(aNodeType(iCounter))="bodies" Then
							Set objBodies=objPart.Item(iCount)
						End If
						
						Set objPart=objPart.Item(iCount)

						bFound=True
						Exit For
					End If
				Next
				If bFound=False Then
					Exit For
				End If
			Next

			'Function returns of object
			If bFound=True Then
				Set Fn_CATIA_GetTreeNodeObject=objPart
			End If
			'Releasing object
			Set objPart=Nothing
			Set objBodies=Nothing
			Set objDocument =Nothing
			Set objCATIA=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Hybrid tree node object
		Case "Hybrid"
			'creating object of Active document of CATIA application
			Set objDocument = objCATIA.ActiveDocument

			aNodeType=Split(sTreeNodeTypes,"~")
			aNode=Split(sTreeNodePath,"~")

			'Mapping Actual value with expected values
			If aNode(0)<>Trim(objDocument.Product.Name) and aNode(0)<>Trim(objDocument.Product.PartNumber) Then
				Set objCATIA=Nothing
				Exit Function
			End If

			If ubound(aNode)=0 Then
				Set Fn_CATIA_GetTreeNodeObject=objDocument.Product
				Set objCATIA=Nothing
				Exit Function
			End If
			'Creating object of all products
			Set objProducts= objDocument.Product.Products
			'Storing Last node name
			sLastNode= aNode(Ubound(aNode))		

			For iCounter=1 to Ubound(aNode)			
				iNextStartCounter=iCounter+1
				If lcase(aNodeType(iCounter))="product" Then
					bFound=False
					For iCount=1 to objProducts.Count
						If Trim(objProducts.Item(iCount).Name)=aNode(iCounter) or Trim(objProducts.Item(iCount).PartNumber)=aNode(iCounter) Then
							If lcase(aNodeType(iCounter+1))<>"product" Then
								Set objProducts=objProducts.Item(iCount)
							Else
								Set objProducts=objProducts.Item(iCount).Products
							End If
							bFound=True
							Exit For
						End If
						sCurrentNode=aNode(iCounter)
					Next
					If bFound=False Then
						Set objProducts=Nothing
						Set objCATIA=Nothing
						Exit Function
					End If
				Else
					bFound=False
					If lcase(aNodeType(iCounter))="part" Then
						aPartName=Split(aNode(iCounter),"_")
						'Matching part with current part
						If Instr(1,objProducts.ReferenceProduct.Parent.Name,aNode(iCounter)) Then
							Set objPart=objProducts.ReferenceProduct.Parent.Part
						ElseIf Instr(1,objProducts.ReferenceProduct.Parent.Name,aPartName(0)) Then
							Set objPart=objProducts.ReferenceProduct.Parent.Part
						Else
							Set objProducts=Nothing
							Set objCATIA=Nothing
							Exit Function
						End If
						sCurrentNode=aNode(iCounter)
						If Lcase(aNodeType(ubound(aNodeType)))="part" Then
							Set Fn_CATIA_GetTreeNodeObject=objPart
							Set objPart=Nothing
							Set objProducts=Nothing
							Set objCATIA=Nothing
							Exit Function
						End If
					End If
					If  lcase(aNodeType(iCounter))="relations" Then
						For iCount = 1 to objDocument.Product.Relations.count 
							bFound =False
							If Instr(1,objDocument.Product.Relations.Item(iCount).Name, aNode(iCounter+1) ) Then
								Set objPublish =objDocument.Product.Relations.Item(iCount)
								bFound = True
								Exit For
							End If
							sCurrentNode=aNode(iCounter+1)
						Next

						If bFound = False Then
							Set objProducts = Nothing
							Set objCATIA = Nothing
							Exit Function
						End If

						If Lcase(aNodeType(ubound(aNodeType)))="relationname" Then
							Set Fn_CATIA_GetTreeNodeObject=objPublish
							Set objPublish=Nothing
							Set objProducts=Nothing
							Set objCATIA=Nothing
							Exit Function
						End If
					End If

					If  lcase(aNodeType(iCounter))="publications"Then
						For iCount = 1 to objProducts.ReferenceProduct.Publications.count 
							bFound =False
							If Instr(1,objProducts.ReferenceProduct.Publications.Item(iCount).Name, aNode(iCounter+1) ) Then
								Set objPublish = objProducts.ReferenceProduct.Publications.Item(iCount).Valuation
								bFound = True
								Exit For
							End If
							sCurrentNode=aNode(iCounter+1)
						Next

						If bFound = False Then
							Set objProducts = Nothing
							Set objCATIA = Nothing
							Exit Function
						End If

						If Lcase(aNodeType(ubound(aNodeType)))="publicationsubpart" Then
							Set Fn_CATIA_GetTreeNodeObject=objPublish
							Set objPublish=Nothing
							Set objProducts=Nothing
							Set objCATIA=Nothing
							Exit Function
						End If
					End If
					
					If sLastNode="xy plane" or sLastNode="zx plane" or sLastNode="yz plane" Then
						Select Case sLastNode
							Case "xy plane"
								Set objPart=objPart.OriginElements.PlaneXY
							Case "zx plane"
								Set objPart=objPart.OriginElements.PlaneZX
							Case "yz plane"
								Set objPart=objPart.OriginElements.PlaneYZ
						End Select
						Set Fn_CATIA_GetTreeNodeObject=objPart
						Set objPart=Nothing
						Set objProducts=Nothing
						Set objCATIA=Nothing
						Exit Function
					End If

					For iCounter1=iNextStartCounter to Ubound(aNode)
						sCurrentNode=aNode(iCounter1)
						bFound=False
						Select Case Lcase(aNodeType(iCounter1))
							Case "bodies"
								Set objPart=objPart.Bodies
							Case "shapes"
								Set objPart=objBodies
								Set objPart=objPart.Shapes
							Case "sketches"
								Set objPart=objBodies
								Set objPart=objPart.Sketches
							Case "geometricelements"
								Set objPart=objPart.GeometricElements
							Case "constraints"
								Set objPart=objPart.Constraints
							Case "originelements"
								'NA
							Case "externalreferences"
								Set objPart= objPart.HybridBodies
							Case "surface"
								Set objPart = objPart.HybridShapes
							Case "material"
								Set objPart = objPart.Parameters
						End Select

						For iCount1=1 to objPart.Count
							If Trim(objPart.Item(iCount1).Name)=aNode(iCounter1) Then
								If Lcase(aNodeType(iCounter1))="bodies" Then
									Set objBodies=objPart.Item(iCount1)
								End If
								Set objPart=objPart.Item(iCount1)
								bFound=True
								Exit For
							End If
						Next

						If bFound=False Then
							Set objPart=Nothing
							Set objProducts=Nothing
							Set objCATIA=Nothing
							Exit Function
						End If
					Next
				End If
				If sCurrentNode=sLastNode Then
					Exit For
				End If
			Next
	
			'Function returns of object
			If bFound=True Then
				If Lcase(aNodeType(ubound(aNodeType)))="product" Then
					Set Fn_CATIA_GetTreeNodeObject=objProducts
				Elseif Lcase(aNodeType(ubound(aNodeType)))="part" Then
					Set Fn_CATIA_GetTreeNodeObject=objPart
				Else
					Set Fn_CATIA_GetTreeNodeObject=objPart
				End If
			End If
			'Releasing object
			Set objPart=Nothing
			Set objProducts=Nothing
			Set objDocument = Nothing
			Set objCATIA=Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get drawing tree node object
		Case "Drawing"	
			'creating object of Active document of CATIA application
			Set objDocument = objCATIA.ActiveDocument
			
			aNode=Split(sTreeNodePath,"~")
			'Mapping Actual value with expected values
			If Instr(1,Trim(objDocument.Name), Trim(aNode(0))) = False Then
				Set objCATIA=Nothing
				Exit Function
			End If
			
			If ubound(aNode)=0 Then
				Set Fn_CATIA_GetTreeNodeObject=objDocument.DrawingRoot
				Set objCATIA=Nothing
				Exit Function
			End If

			Set objSheets= objDocument.DrawingRoot.Sheets
			Set objViews = Nothing

			For iCounter = 1 to UBound(aNode)
				bFound = False
				For iSheetCount = 1 to objSheets.count
					If Trim(objSheets.Item(iSheetCount).Name) = Trim(aNode(iCounter)) Then
						If Ubound(aNode) = iCounter Then
							Set Fn_CATIA_GetTreeNodeObject = objSheets.Item(iSheetCount)
						Else
							Set objViews = objSheets.Item(iSheetCount).Views
						End If
						bFound = True
						Exit For
					End If
				Next
				If bFound = True Then
					Exit For
				End If
			Next

			If objViews Is Nothing = False Then
				For iCounter=1 to Ubound(aNode)
					bFound=False
					For iCount = 1 to objViews.count
						If Trim(objViews.Item(iCount).Name) = aNode(iCounter) Then
							Set Fn_CATIA_GetTreeNodeObject = objViews.Item(iCount)
							bFound = True
							Exit For
						End If
					Next
					If bFlag = True Then
						Exit For
					End If
				Next
			End If

			'Function returns of object
			If bFound=False Then
				Set Fn_CATIA_GetTreeNodeObject=Nothing
			End If
			'Releasing object
			Set objViews=Nothing
			Set objSheets = Nothing
			Set objCATIA=Nothing
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation
'
'Function Description	 :	Function used to perform operations on Teamcenter Save Manager Navigation Tree
'
'Function Parameters	 :  1.sAction 		: Action to perform on tree					
'							2.sNodeName		: Node name
'
'Function Return Value	 : 	True Or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Teamcenter Save Manager Navigation Tree node should be available
'							
'Function Usage		     :	bReturn=Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation("Select","Home~AutomatedTest")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  28-Apr-2017	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation(sAction,sNodeName)
	'Declaring variables
	Dim objNavTree
	Dim aNodeName
	Dim iCounter,iCount
	Dim sParentPath
	Dim bFlag
	
	'Initially function returns false
	Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation=False
	
	'creating object of [ NavTree ]
	Set objNavTree = Fn_FSOUtil_XMLFileOperations("getobject","CATIA_Common_OR","jtree_TeamcenterNavTree","")

	'Checking existance of [ NavTree ]
	If Fn_UI_Object_Operations("Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation","Exist", objNavTree, "","","") = False Then
		Exit Function
	End If

	Select Case sAction
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to select node from nav tree
		Case "Select"
			'Initial Item parent Path
			aNodeName = Split (sNodeName, "~")
			For iCounter =0 to UBound(aNodeName)-1
				If sParentPath = "" Then
					sParentPath  = aNodeName(iCounter)
				Else
					sParentPath  = sParentPath & "~" & aNodeName(iCounter)
				End If
			Next
			
			'Expanding parent node
			If UBound(aNodeName) > 0 Then
				If Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation("Expand",sParentPath) = False Then
					Exit Function
				End If
				wait 1
			End If
			
			'Checking Exsitence of Node
			bFlag = Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation("Exist", sNodeName)
			If bFlag=False Then
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = False
			Else
				'Selecting node from tree
				objNavTree.Select sNodeName
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = True
			End If			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Expand node from nav tree
		Case "Expand"		
			'Checking Exsitence of Node
			bFlag = Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation("Exist", sNodeName)
			If bFlag=False Then
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = False
			Else
				'Expanding node from nav tree
				objNavTree.Expand sNodeName
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = True
			End If			
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Collapse node from nav tree
		Case "Collapse"		
			'Checking Exsitence of Node
			bFlag = Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation("Exist", sNodeName)
			If bFlag=False Then
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = False
			Else
				'Collapse node from nav tree
				objNavTree.Collapse sNodeName
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = True
			End If
		 '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		 'Case to check existance of node of nav tree
		 Case "Exist"
			aNodeName = Split (sNodeName, "~")
			For iCounter =0 to UBound(aNodeName)
				If sParentPath = "" Then
					sParentPath  = aNodeName(iCounter)
				Else
					sParentPath  = sParentPath & "~" & aNodeName(iCounter)
				End If
			
				bFlag=False
				For iCount=0 to objNavTree.GetROProperty("items count")-1
					If Trim(objNavTree.GetItem(iCount))=Trim(sParentPath) Then
						If iCounter<>UBound(aNodeName) Then
							objNavTree.Expand sParentPath
						End If
						wait 1
						bFlag=True
						Exit For
					End If
				Next
			
				If bFlag=False Then
					Exit for
				End If
			Next				
			If bFlag=False Then
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = False
			Else
				Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = True
			End If			
	End Select

	If Err.Number < 0 Then
		Fn_CATIA_TeamcenterSaveManagerNavigationTreeOperation = False
	End If
	
	'Releasing Nav tree object
	Set objNavTree=Nothing	
End Function
