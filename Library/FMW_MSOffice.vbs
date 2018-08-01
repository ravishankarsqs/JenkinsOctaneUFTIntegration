Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name								|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_MSO_ExcelOperations							|	sandeep.navghane@sqs.com	|	16-Jan-2015	|	Function used to perform operation on MS excel file
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_MSO_ExcelOperations
'
'Function Description	 :	Function used to perform operation on MS excel file
'
'Function Parameters	 :  1.Action			: Action name to perform
'							2.sExcelFilePath	: Excel file path
'							3.iSheetNumber		: Excel Sheet number on which user wants to perform operations
'							4.sCellPosition		: Cell position
'							5.sValue			: text value
'							6.bCloseExcel		: Excel file close option
'
'Function Return Value	 : 	True \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Excel file should exist
'
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("GetCellRowAndColumnNumberPosition","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,"","Test Case Name",True)
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("GetCellRowNumberOfSpecificColumn","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,3,"FAIL",True)
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("Autofilter","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,"Result","PASS",True)
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("Sort","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,"Test Case Name","",True)
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("GetColumnData","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,"Test Case Name","",True)
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("GetColumnNames","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,"","",True)
'Function Usage		     :  bReturn = Fn_MSO_ExcelOperations("GetColumnNames","C:\GOG_AL_11.2\Reports\BtachExecutionDetails.xlsx",1,2,"",True)
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_MSO_ExcelOperations(sAction,sExcelFilePath,iSheetNumber,sCellPosition,sValue,bCloseExcel)
	Err.Clear
	'Declaring variables
	Dim objExcel,objWorkbook,objWorksheet,objRange,objRange2
	Dim iRowNumber,iColumnNumber,iLastColumn,iLastRow,iFilterColumnNumber,iCounter
	Dim aCellPosition
	Dim bFlag
	Dim sTempValue
	
	Const xlCellTypeLastCell = 11
	Const xlAscending = 1
	Const xlYes = 1
	'Initially set function return value as False
	Fn_MSO_ExcelOperations = False
	
	'Creating excel file object	
	If sExcelFilePath<>"" Then
		Set objExcel = CreateObject("Excel.Application")
		Set objWorkbook = objExcel.Workbooks.Open(sExcelFilePath)
		objExcel.Visible = false
		objExcel.DisplayAlerts = False
	Else
		Set objExcel =GetObject("","Excel.Application")
		Set objWorkbook = objExcel.Workbooks(1)
	End If
	If iSheetNumber="" Then
		iSheetNumber=1
	End If
	
	Select case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get row and column position of specific text in excel
		Case "GetCellRowAndColumnNumberPosition"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			'Activating worksheet
			objWorksheet.Activate
			objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
			For iRowNumber = 1 To objExcel.ActiveCell.Row
				For iColumnNumber = 1 to objExcel.ActiveCell.Column
					If LCase(objExcel.Cells(iRowNumber,iColumnNumber).Value) = LCase(sValue) Then
						Fn_MSO_ExcelOperations = iRowNumber & ":" & iColumnNumber
						Exit For
					End If
				Next
			Next
			'Releasing worksheet object
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get row number of specific text in excel under specific column (passed by user)
		Case "GetCellRowNumberOfSpecificColumn"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			'Activating worksheet
			objWorksheet.Activate
			objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
			For iRowNumber = 1 To objExcel.ActiveCell.Row
				If LCase(objExcel.Cells(iRowNumber,sCellPosition).Value) = LCase(sValue) Then
					Fn_MSO_ExcelOperations = iRowNumber
					Exit For
				End If
			Next
			'Releasing worksheet object
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to set specific cell data
		Case "SetCellData"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			'Activating worksheet
			objWorksheet.Activate
			aCellPosition=Split(sCellPosition,":")
			objExcel.Cells(Cint(aCellPosition(0)),Cint(aCellPosition(1))).Value=sValue
			Fn_MSO_ExcelOperations = True
			'Releasing worksheet object
			objWorksheet.SaveAs sExcelFilePath
			'Releasing worksheet object
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get specific cell data
		Case "GetCellData"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			'Activating worksheet
			objWorksheet.Activate
			aCellPosition=Split(sCellPosition,":")
			Fn_MSO_ExcelOperations=objExcel.Cells(Cint(aCellPosition(0)),Cint(aCellPosition(1))).Value
			'Releasing worksheet object
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to filter data
		Case "Autofilter"
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			objWorksheet.Activate
			iRowNumber=-1
			Set objRange = objWorkSheet.UsedRange		
			iLastRow = objRange.Row + objRange.Rows.Count - 1 
			iLastColumn = objRange.Column + objRange.Columns.Count - 1
			bFlag=False
			For iRowNumber = 1 To iLastRow
				For iColumnNumber = 1 to objRange.Columns.Count
					If Not IsEmpty(objExcel.Cells(iRowNumber,iColumnNumber).Value) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=True Then
					Exit For
				End If
			Next

			If iRowNumber=-1 Then
				Exit Function
			End If
			iFilterColumnNumber=-1
			For iColumnNumber = 1 to iLastColumn
				If LCase(objExcel.Cells(iRowNumber,iColumnNumber).Value) = LCase(sCellPosition) Then
					iFilterColumnNumber=iColumnNumber
					Exit For
				End If
			Next
			If iFilterColumnNumber=-1 Then
				Exit Function
			End If

			objWorksheet.Cells(iRowNumber,iLastColumn).autofilter iFilterColumnNumber,sValue
			objExcel.ActiveWorkbook.Save
			Fn_MSO_ExcelOperations=True
			Set objRange =Nothing
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to Sort data
		Case "Sort"		
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			objWorksheet.Activate
			iRowNumber=-1
			Set objRange = objWorkSheet.UsedRange		
			iLastRow = objRange.Row + objRange.Rows.Count - 1 
			iLastColumn = objRange.Column + objRange.Columns.Count - 1
			bFlag=False
			For iRowNumber = 1 To iLastRow
				For iColumnNumber = 1 to objRange.Columns.Count
					If Not IsEmpty(objExcel.Cells(iRowNumber,iColumnNumber).Value) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=True Then
					Exit For
				End If
			Next

			If iRowNumber=-1 Then
				Exit Function
			End If

			iFilterColumnNumber=-1
			For iColumnNumber = 1 to iLastColumn
				If LCase(objExcel.Cells(iRowNumber,iColumnNumber).Value) = LCase(sCellPosition) Then
					iFilterColumnNumber=iColumnNumber
					Exit For
				End If
			Next
			If iFilterColumnNumber=-1 Then
				Exit Function
			End If
				
			Set objRange2 = objExcel.Cells(iRowNumber,iFilterColumnNumber)
			If sValue="" or sValue="Ascending" Then
				objRange.Sort objRange2, xlAscending, , , , , , xlYes
			End IF

			objExcel.ActiveWorkbook.Save
			Fn_MSO_ExcelOperations=True
			Set objRange2 =Nothing
			Set objRange =Nothing
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get specific column data
		Case "GetColumnData"		
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			objWorksheet.Activate
			iRowNumber=-1
			Set objRange = objWorkSheet.UsedRange		
			iLastRow = objRange.Row + objRange.Rows.Count - 1 
			iLastColumn = objRange.Column + objRange.Columns.Count - 1
			bFlag=False
			For iRowNumber = 1 To iLastRow
				For iColumnNumber = 1 to objRange.Columns.Count
					If Not IsEmpty(objExcel.Cells(iRowNumber,iColumnNumber).Value) Then
						bFlag=True
						Exit For
					End If
				Next
				If bFlag=True Then
					Exit For
				End If
			Next

			If iRowNumber=-1 Then
				Exit Function
			End If

			iFilterColumnNumber=-1
			For iColumnNumber = 1 to iLastColumn
				If LCase(objExcel.Cells(iRowNumber,iColumnNumber).Value) = LCase(sCellPosition) Then
					iFilterColumnNumber=iColumnNumber
					Exit For
				End If
			Next
			If iFilterColumnNumber=-1 Then
				Exit Function
			End If
				
			iRowNumber=iRowNumber+1
			sTempValue=""
			For iCounter = iRowNumber To iLastColumn
				If Not IsEmpty(objExcel.Cells(iCounter,iFilterColumnNumber).Value) Then
					If sTempValue="" Then
						sTempValue=objExcel.Cells(iCounter,iFilterColumnNumber).Value
					Else
						sTempValue=sTempValue & "~" & objExcel.Cells(iCounter,iFilterColumnNumber).Value	
					End If
				End IF
			Next
			If sTempValue<>"" Then
				Fn_MSO_ExcelOperations=sTempValue	
			End If			
			Set objRange =Nothing
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to to find value in excel
		Case "FindValue"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			objWorkSheet.Activate
			'Creating used range object		
			Set objRange = objWorkSheet.UsedRange.Find(Cstr(sValue))
			If TypeName(objRange)<>"Nothing" Then
				If Cstr(objRange)=Cstr(sValue) Then
					Fn_MSO_ExcelOperations = True
				End If
			End If
			Set objRange =Nothing
			'Releasing worksheet object
			Set objWorksheet = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get all column names
		Case "GetColumnNames"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			'Activating worksheet
			objWorksheet.Activate
			objWorksheet.UsedRange.SpecialCells(xlCellTypeLastCell).Activate
			bFlag=False
			If sCellPosition="" Then
				For iRowNumber = 1 To objExcel.ActiveCell.Row	
					For iColumnNumber = 1 to objExcel.ActiveCell.Column
						If Not IsEmpty(objExcel.Cells(iRowNumber,iColumnNumber).Value) Then
							bFlag=True
							Exit For
						End If
					Next
					If bFlag=True Then
						Exit For
					End If
				Next
			Else
				iRowNumber=Cint(sCellPosition)
				bFlag=True
			End If
			
			sTempValue=""
			If bFlag=True Then
				For iColumnNumber = 1 to objExcel.ActiveCell.Column
					If sTempValue="" Then
						sTempValue=objExcel.Cells(iRowNumber,iColumnNumber).Value
					Else
						sTempValue=sTempValue & "~" & objExcel.Cells(iRowNumber,iColumnNumber).Value
					End If
				Next
			End If
			If Err.Number <> 0 Then
				Fn_MSO_ExcelOperations=""
				Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_MSO_ExcelOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
			Else
				Fn_MSO_ExcelOperations=sTempValue
			End If
			'Releasing worksheet object
			Set objWorksheet = Nothing
			Exit Function
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to copy cell data
		Case "CopyCell"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			aCellPosition=Split(sCellPosition,":")
'			Extern.EmptyClipboard()
			Fn_MSO_ExcelOperations=objWorksheet.Cells(CInt(aCellPosition(0)),CInt(aCellPosition(1))).Value
'			objWorksheet.Copy
			Set objWorksheet =Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to paste data in cell
		Case "PasteSpecial"
			'Creating worksheet object
			Set objWorksheet = objWorkbook.Worksheets(iSheetNumber)
			aCellPosition=Split(sCellPosition,":")
			objWorksheet.Cells(CInt(aCellPosition(0)),CInt(aCellPosition(1))).Value=sValue
			'objWorksheet.range(sCellPosition).PasteSpecial
'			objWorksheet.Range(sCellPosition).Paste
'			objWorksheet.Range(sCellPosition).Select
'			objWorksheet.Paste
			Set objWorksheet =Nothing
	End Select
	
	'Closing excel file
	Select Case Cbool(bCloseExcel)
		Case True
			objExcel.DisplayAlerts = False
			objExcel.Quit
			wait 2
	End Select
	
	If Err.Number <> 0 Then
		Fn_MSO_ExcelOperations=False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_MSO_ExcelOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
	
	'Releasing excel file object
	Set objWorkbook = Nothing
	Set objExcel = Nothing	
End Function