'! @Name 			RAC_Common_TabFolderWidgetOperations
'! @Details 		This actionword is used to perform operations on teamcenter tabs
'! @InputParam1 	sAction : Action to be performed
'! @InputParam2 	sItem 	: Tab item name
'! @InputParam3 	sMenu 	: Popup menu name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			23 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","Overview",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"DoubleClick","Summary",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"RMBMenuSelect","Detail","Close"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Close","Overview",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"VerifyActivate","Summary",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"VerifyExist","Results",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"GetTabToolTipText","Summary",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sItem,sPopupMenu
Dim iXLen,iYLen,iXBound,iCounter,iTabItemCount,iTabIndex,iIndexCounter
Dim objRACTabFolderWidget,objItem,objContextMenu
Dim aBounds,aItem
Dim sBounds
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action input parameters in local variables
sAction = Parameter("sAction")
sItem = Parameter("sItem")
sPopupMenu = Parameter("sPopupMenu")

aItem = Split(sItem,"@")

'Fetching datatable current selected row
GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Creating object of [ RACTabFolderWirdget ] object
Set objRACTabFolderWidget=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jobj_RACTabFolderWidget","")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_TabFolderWidgetOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

objRACTabFolderWidget.SetTOProperty "Index", 0
iTabItemCount = objRACTabFolderWidget.Object.getTabItemCount

bFlag = False
For iIndexCounter = 0 to 5
	objRACTabFolderWidget.SetTOProperty "Index",iIndexCounter
	iXLen = 0
	iYLen = 0
	If uBound(aItem) > 0 Then
		iTabIndex = cInt(aItem(1)) - 1
	Else
		iTabIndex = 0
	End If
	'Checking existance of [ RACTabFolderWirdget ] object
	If objRACTabFolderWidget.Exist(2) Then
		'getting tab item count
		iTabItemCount = objRACTabFolderWidget.Object.getTabItemCount
		For iCounter = 0 to iTabItemCount-1
			'Creating object of specific tab item
			Set objItem = objRACTabFolderWidget.Object.getItem(iCounter)
			iXLen = iXLen + objItem.getWidth
			If Trim(objItem.text) = Trim(aItem(0)) Then
				If iTabIndex = 0 Then
					bFlag = True
					Exit for
				Else
					iTabIndex = iTabIndex - 1
				End If
			End If
		Next
		If bFlag = True Then
			Exit for
		End If
	End If
Next

If bFlag = False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] as tab [ " & Cstr(aItem(0)) & " ] does not exist","","","","","")
	Call Fn_ExitTest()
Else
	sItem=aItem(0)
End If

If sAction="Select" Then
	bFlag=objItem.isShowing()
	If bFlag="false" Then
		'double clicking on tab object to visible all tab items
		objRACTabFolderWidget.DblClick 1,1,"LEFT"
		wait GBL_MIN_MICRO_TIMEOUT
	End If
End If

'getting tab bounds
sBounds = objItem.getBounds().toString()
sBounds = mid(sBounds,instr(sBounds,"{")+1, len(sBounds) -instr(sBounds,"{")-1)
aBounds = split(sBounds,",")
iXBound = cInt(Trim(aBounds(0)))
iXLen = iXBound + 15
iYLen = (objItem.getHeight/2)           

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
	
Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select specific tab item
	Case "Select"                                                     
		objRACTabFolderWidget.Click iXLen,iYLen,"LEFT"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		If bFlag="false" Then
			objRACTabFolderWidget.DblClick iXLen,iYLen,"LEFT"
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		End if		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select [ " & Cstr(sItem) & " ] tab","","","","","")
			Call Fn_ExitTest()
		End If
		'Capturing execution execution time
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected [ " & Cstr(sItem) & " ] tab","","","","DONOTSYNC","")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to double click on specific tab item
	Case "DoubleClick" 
		objRACTabFolderWidget.DblClick iXLen, iYLen, "LEFT"
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to double click on [ " & Cstr(sItem) & " ] tab","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully double click on [ " & Cstr(sItem) & " ] tab","","","","DONOTSYNC","")		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to select RMB menu of specific tab item
	Case "RMBMenuSelect"
		'Right click on tab item
		objRACTabFolderWidget.Click iXLen, iYLen, "RIGHT"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		aMenuList = Split(sPopupMenu,":",-1,1)
		'Creating object of [ Context menu ] object
		Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","wmnu_ContextMenu","")

		'Select Menu action
		Select Case Ubound(aMenuList)
			Case "0"
				sPopupMenu = objContextMenu.BuildMenuPath(aMenuList(0))
			Case "1"
				sPopupMenu = objContextMenu.BuildMenuPath(aMenuList(0),aMenuList(1))
			Case "2"
				sPopupMenu = objContextMenu.BuildMenuPath(aMenuList(0),aMenuList(1),aMenuList(2))
			Case Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu as RMB menu [ " & Cstr(sPopupMenu) & " ] does not exist for [ " & Cstr(sItem) & " ] tab","","","","","")
				Set objContextMenu=Nothing	
				Call Fn_ExitTest()
		End Select
		
		'Selecting RMB menu
		objContextMenu.Select sPopupMenu
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select RMB menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sItem) & " ] tab","","","","","")
			Set objContextMenu=Nothing
			Call Fn_ExitTest()
		End If
		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_MICRO_SYNC_ITERATIONS)
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected RMB menu [ " & Cstr(sPopupMenu) & " ] on [ " & Cstr(sItem) & " ] tab","","","","","")
		'Releasing object of [ Context menu ] object
		Set objContextMenu=Nothing
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify specific tab item is in active state
	Case "VerifyActivate"
		iCounter = objRACTabFolderWidget.Object.getSelectedTabIndex
		Set objItem = objRACTabFolderWidget.Object.getItem(iCounter)
		
		If Trim(objItem.text) = Trim(sItem) Then
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified [ " & Cstr(sItem) & " ] tab is currently activated","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as [ " & Cstr(sItem) & " ] tab is not activated","","","","","")
			Call Fn_ExitTest()
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify specific tab item is available
	Case "VerifyExist"
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Sucessfully verified [ " & Cstr(sItem) & " ] tab exist\available","","","","DONOTSYNC","")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to close specific tab item
	Case "Close"
		objRACTabFolderWidget.Click iXLen, iYLen, "LEFT"
		If sItem="Properties" Then
			'do nothing
		Else
			Call Fn_RAC_ReadyStatusSync(2)
		End If
		
		sBounds = objItem.getCloseButtonBounds.toString()
		sBounds = right(sBounds, Len(sBounds)-instr(sBounds, "{"))
		aBounds = split(sBounds, ",", -1, 1)
		iXLen = Cint(Trim(aBounds(0))) + 5
		iYLen = Cint(Trim(aBounds(1))) + 5
		objRACTabFolderWidget.Click iXLen, iYLen,"LEFT"
		
		If Err.Number <> 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to close [ " & Cstr(sItem) & " ] tab","","","","","")
			Call Fn_ExitTest()
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Teamcenter Tab Operations",sAction,"Tab Name",sItem)
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully closed [ " & Cstr(sItem) & " ] tab","","","","","")
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to tool tip test of specific tab item
	Case "GetTabToolTipText"
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
		Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
		DataTable.SetCurrentRow 1		
		DataTable.Value("ReusableActionWordName","Global")= "RAC_Common_TabFolderWidgetOperations"
		DataTable.Value("ReusableActionWordReturnValue","Global")= objItem.getToolTipText()
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case Else
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform any operation on teamcenter tab due to invalid case","","","","","")
		Call Fn_ExitTest()
End Select

'Set tab object index to 1
objRACTabFolderWidget.SetTOProperty "Index", 1

'Releasing all objects
Set objItem = Nothing
Set objRACTabFolderWidget = Nothing

Function Fn_ExitTest()
	'Releasing all objects
	Set objItem = Nothing
	Set objRACTabFolderWidget = Nothing
	ExitTest
End Function


