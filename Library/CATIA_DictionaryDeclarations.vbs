Option Explicit

'Declaring variables
Dim dictCircleInfo
Dim dictRectangleInfo
Dim dictExtrudeInfo
Dim dictCATIATCSaveManagerInfo
Dim dictPadInfo
Dim dictDrawingViewsInfo
Dim dictCATIANavigaionTreeInfo
Dim dictCATIAImportSpreadSheetInfo

'to store information required to create circle in CATIA
Set dictCircleInfo = CreateObject("Scripting.Dictionary")

'to store information required to create extrude in CATIA
Set dictExtrudeInfo = CreateObject("Scripting.Dictionary")

'to store information required to perform TC save manager related operations in CATIA
Set dictCATIATCSaveManagerInfo = CreateObject("Scripting.Dictionary")

'to store information required to create rectangle in CATIA
Set dictRectangleInfo = CreateObject("Scripting.Dictionary")

'to store information required to pad shapes in CATIA
Set dictPadInfo = CreateObject("Scripting.Dictionary")

'to store Drawing Views information in CATIA
Set dictDrawingViewsInfo=CreateObject("Scripting.Dictionary")

'to store Drawing Views information in CATIA
Set dictCATIANavigaionTreeInfo=CreateObject("Scripting.Dictionary")

'to store information required to updatespread sheet in CATIA
Set dictCATIAImportSpreadSheetInfo = CreateObject("Scripting.Dictionary")
