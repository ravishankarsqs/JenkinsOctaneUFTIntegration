Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'		Function Name									|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. 	Fn_NX_ReadyStatusSync							|	sandeep.navghane@sqs.com	|	28-Mar-2016	|	Function used to waits till NX Application comes to Ready state
'002. 	Fn_NX_RestoreNXMainWindow						|	sandeep.navghane@sqs.com	|	28-Mar-2016	|	Function used to restore NX application main window
'003. 	Fn_NX_SetCursorToStandardPosition				|	sandeep.navghane@sqs.com	|	28-Mar-2016	|	Function used to set cursor at standard location\postion in NX application
'004. 	Fn_NX_GetTreeNodePath							|	sandeep.navghane@sqs.com	|	30-Mar-2016	|	Function used to retrive tree node path from exported values
'005. 	Fn_NX_GetMenuRealPath							|	sandeep.navghane@sqs.com	|	26-May-2016	|	Function used to get real NX menu path
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_NX_ReadyStatusSync
'
'Function Description	 :	Function used to waits till NX Application comes to Ready state
'
'Function Parameters	 :  1.iIterations: No. of times to be checked for Ready text						
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NX application should be displayed
'
'Function Usage		     :	Call Fn_NX_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  26-Feb-2016	    |	 1.0		|	  Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_NX_ReadyStatusSync(iIterations)
	'Declaring Variables
	Dim iCounter, iCount
	Dim objProgressBar
	
	'Creating object of NX Window progrss bar
	Set objProgressBar =Window("text:=","regexpwndclass:=Afx:","is owned window:=False","is child window:=False").WinObject("regexpwndclass:=SysAnimate32")
	
	For iCounter = 1 to iIterations
		For iCount=1 to 6
			If Fn_WIN_UI_WinObject_Operations("Fn_NX_ReadyStatusSync","Exist",objProgressBar,"1","","") Then
				objProgressBar.WaitProperty "exist",true,10000
			Else
				Wait 0,350
				Exit For
			End If
		Next
	Next
	
	'Handling cursor state
	If Fn_NX_SetCursorToStandardPosition()=True Then
		For iCounter = 1 to iIterations
			For iCount=1 to 24
				If Fn_CommonUtil_GetCursorState()="65539" Then
					Wait 0,300
					Exit For
				Else
					wait GBL_MICRO_TIMEOUT
				End IF
			Next
		Next
	End If
	
	If Fn_WIN_UI_WinObject_Operations("Fn_NX_ReadyStatusSync","Exist",objProgressBar,GBL_MIN_TIMEOUT+1,"","") = False Then
		Fn_NX_ReadyStatusSync = True
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<PASS>: [ Fn_NX_ReadyStatusSync ] : NX is Ready in [ " & CStr(iIterations) & " ] sync iterations")		
	Else
		Fn_NX_ReadyStatusSync = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"), "<FAIL>: [ Fn_NX_ReadyStatusSync ] : NX Not Ready after [ " & CStr(iIterations) & " ] sync iterations")		
	End If
	
	'Release Object
	Set objProgressBar = Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_NX_RestoreNXMainWindow
'
'Function Description	 :	Function used to restore NX application main window
'
'Function Parameters	 :  NA
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NX application should be available
'
'Function Usage		     :	Call Fn_NX_RestoreNXMainWindow()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  28-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_NX_RestoreNXMainWindow()
	On Error Resume Next
	'Declaring Variables
	Dim objNXMainWindow	
	'Creating object of NX Window
	Set objNXMainWindow=Window("regexpwndtitle:=NX .*","regexpwndclass:=Afx:")
	
	If objNXMainWindow.Exist(1)  Then
		If objNXMainWindow.GetROProperty("visible")=False Then
			objNXMainWindow.Restore
		End If
	End If
	'Releasing object of NX Window
	Set objNXMainWindow =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_NX_SetCursorToStandardPosition
'
'Function Description	 :	Function used to set cursor at standard location\postion in NX application
'
'Function Parameters	 : 	NA
'
'Function Return Value	 :  NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NX application should be available
'
'Function Usage		     :	Call Fn_NX_SetCursorToStandardPosition()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  28-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_NX_SetCursorToStandardPosition()	
	'Declaring Variables
	Dim objNXMainWindow	

	Fn_NX_SetCursorToStandardPosition=False
	
	'Creating object of NX Window
	Set objNXMainWindow=Window("regexpwndtitle:=NX .*","regexpwndclass:=Afx:")
	If CBool(objNXMainWindow.GetROProperty("enabled"))=True Then
		'Restore NX application
		Call Fn_NX_RestoreNXMainWindow()
		wait GBL_MICRO_TIMEOUT
		objNXMainWindow.Click Cint(objNXMainWindow.getroproperty("width")/2),(objNXMainWindow.getroproperty("height")-15)
		Fn_NX_SetCursorToStandardPosition=True
	End If
	'Releasing object of NX Window
	Set objNXMainWindow =Nothing
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_NX_GetTreeNodePath
'
'Function Description	 :	Function used to retrive tree node path from exported values
'
'Function Parameters	 :  1.sExportOption 	: Tree exported option
'							2.sNodePath	  		: Tree node
'							3.sDelimiter 		: Tree node delimiter
'							4.sInstanceHandler 	: Tree node Instance Handler
'
'Function Return Value	 : 	False or Tree node path
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Tree node should be available
'
'Function Usage		     :	bReturn=Fn_NX_GetTreeNodePath("Excel","Groups~Title Group~Title Block1","~","@")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  30-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_NX_GetTreeNodePath(sExportOption,sNodePath,sDelimiter,sInstanceHandler)
	'Declaring variables
	Dim objExcel,objWorkSheet
	Dim iTempStartRowNumber,iCounter,iNodePath,iRowNumber,iStart,iPrevSpaceCount,iCount, iSpaceCount 
	Dim bFlag
	Dim aNodePath
	Dim sTempString
	
	Fn_NX_GetTreeNodePath=False
	
	Select Case sExportOption
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		'Case to retrive node path from excel file
		Case "Excel"
			'Creating excel file object
			Set objExcel =GetObject(, "Excel.Application")
			aNodePath=Split(sNodePath,"~")
			iTempStartRowNumber=2
			objExcel.Worksheets(1).Activate
			'Creating object of work sheet
			Set objWorkSheet=objExcel.Worksheets(1)
			iPrevSpaceCount=0
			For iCounter=0 to ubound(aNodePath)
				bFlag=False
				iStart=1
				For iRowNumber =iTempStartRowNumber To objWorkSheet.UsedRange.Rows.Count
					iSpaceCount=0
					sTempString=objExcel.Cells(iRowNumber,1).Value
					For iCount=1 to Len(sTempString)
						If mid(sTempString,iCount,1)=" " Then
							iSpaceCount=iSpaceCount+1
						Else
							Exit For
						End If
					Next       
					If Instr(Trim(LCase(objExcel.Cells(iRowNumber,1).Value)), "(order: chronological)") > 0 Then
						objExcel.Cells(iRowNumber,1).Value = Replace(objExcel.Cells(iRowNumber,1).Value, "(Order: Chronological)","")
					End If					
					If Trim(LCase(objExcel.Cells(iRowNumber,1).Value)) = LCase(aNodePath(iCounter)) Then
						bFlag=True
						If iNodePath<>"" Then
							iNodePath=iNodePath & " " & iStart
						Else
							iNodePath=iStart
						End If
						
						iTempStartRowNumber=iRowNumber+1

						sTempString=objExcel.Cells(iTempStartRowNumber,1).Value
						iSpaceCount=0
						For iCount=1 to Len(sTempString)
							If mid(sTempString,iCount,1)=" " Then
								iSpaceCount=iSpaceCount+1
							Else
								Exit For
							End If
						Next
						iPrevSpaceCount=iSpaceCount
						Exit For
					ElseIf iPrevSpaceCount=iSpaceCount Then
						iStart=iStart+1
						iPrevSpaceCount=iSpaceCount
					End If
				Next
				If bFlag=False Then
					Exit For
				End If
			Next
			If bFlag=True Then
				Fn_NX_GetTreeNodePath=Cstr(iNodePath)
			End If
			objExcel.DisplayAlerts = False
			objExcel.Quit
			wait GBL_MIN_TIMEOUT
			'Releasing object of work sheet
			Set objWorkSheet=Nothing
			'Releasing excel file object
			Set objExcel =Nothing
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_NX_GetMenuRealPath
'
'Function Description	 :	Function used to get real NX menu path
'
'Function Parameters	 : 	1.sMenuCaption : Menu caption
'
'Function Return Value	 :  Real menu path
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	Call Fn_NX_GetMenuRealPath("UG_FILE_NEW")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	  Reviewer			|	Changes Done
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  26-May-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_NX_GetMenuRealPath(sMenuCaption)
	Select Case sMenuCaption
		Case "UG_FILE_NEW"
			Fn_NX_GetMenuRealPath="File -> New..."
		Case "UG_FILE_QUIT"
			Fn_NX_GetMenuRealPath="File -> Exit"
		Case "UG_FILE_IMPORT_PART"
			Fn_NX_GetMenuRealPath="File -> Import -> Part..."
		Case "UG_FILE_OPEN"
			Fn_NX_GetMenuRealPath="File -> Open..."
		Case "UG_APP_MODELING"
			Fn_NX_GetMenuRealPath="Application -> Modeling"
		Case "UG_APP_DRAFTING"
			Fn_NX_GetMenuRealPath="Application -> Drafting"
		Case "UG_JOURNAL_PLAY"
			Fn_NX_GetMenuRealPath="Tools -> Journal -> Play"
		Case "UG_FILE_SAVE_AND_CLOSE"
			Fn_NX_GetMenuRealPath="File -> Close -> Save and Close"
		Case Else
			Fn_NX_GetMenuRealPath=sMenuCaption
	End Select
End Function