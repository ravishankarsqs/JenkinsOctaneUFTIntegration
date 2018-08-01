'! @Name 			RAC_Common_CATIASubMenuOperations
'! @Details 		Action word to perform operations on sub menu of CATIA menu in teamcenter
'! @InputParam1 	sAction 				: String to indicate what action is to be performed on structure manager bom table e.g. Select, Expand
'! @InputParam2 	sPerformOperationFrom 	: From where user wants to perform operation on Catia sub menu
'! @InputParam3 	sNodeName 				: Node path
'! @InputParam4 	sMenuLabel 				: Menu label tag
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			18 Nov 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "Select","","","CATIACreateExportSpreadsheet"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CATIASubMenuOperations","RAC_Common_CATIASubMenuOperations", oneIteration, "SelectWithoutSync","","","CATIAReturnToCATIA"

'Declaring variables
Dim sAction,sPerformOperationFrom,sNodeName,sMenuLabel
Dim sCaseName,sMenuLabelTemp,sPerspective
Dim iPerformanceTime,iCounter,iWindowCount,iExpectedWindowCount
Dim objCATIA,objWindows,objInformation,objProcessDialog
Dim bFLag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sAction = Parameter("sAction")
sPerformOperationFrom = Parameter("sPerformOperationFrom")
sNodeName = Parameter("sNodeName")
sMenuLabel = Parameter("sMenuLabel")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CATIASubMenuOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

If sMenuLabel<>"" Then
	'Storing menu label
	sMenuLabelTemp=Fn_RAC_GetXMLNodeValue("RAC_Common_MenuOperations","",sMenuLabel)	
	If sMenuLabelTemp=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to fetch value of menu label [ " & Cstr(sMenuLabel) & " ] from XML while performing menu operation","","","","DONOTSYNC","")
		Call Fn_ExitTest()
	End If
	If sMenuLabelTemp="CATIA V5:Return to CATIA V5" Then
		GBL_CURRENT_EXECUTABLE_APP="CATIA"
	End If
End If

'To Select Node
If sNodeName<>"" Then
	Select Case Lcase(sPerformOperationFrom)
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "mytcnavigationtree"
			LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select",sNodeName,""
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "searchresultstree"
			LoadAndRunAction "RAC_Search\RAC_Search_SearchResultsTreeOperation","RAC_Search_SearchResultsTreeOperation",OneIteration,"Select",sNodeName,"",""
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "pse_bomtable"
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",sNodeName,"","",""
	End Select
	Call Fn_RAC_ReadyStatusSync(DEFAULT_SYNC_ITERATIONS)
End If

'Getting active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

Select Case Lcase(sPerspective)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "structuremanager","myteamcenter"
		Select Case sAction
			Case "Select","SelectExt","SelectWithoutSync"',"SelectVerifyWorkingStatus"
				sCaseName="WinMenuSelectWithoutSync"
				If sMenuLabelTemp="CATIA V5:Load in CATIA V5" Then
					GBL_LOADINCATIAPROCESS_FLAG=True
				End If
			Case "VeirfyExist"
				sCaseName="VeirfyExist"
			Case "ReturnToCATIA"
				sCaseName="WinMenuSelectWithoutSync"
		End Select
End Select

	
Select Case sAction
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Select"
		If sMenuLabelTemp="CATIA V5:Load in CATIA V5" Then
			Set objCATIA=GetObject("","CATIA.Application")
			Set objWindows = objCATIA.Windows
			iWindowCount=objWindows.Count
			iExpectedWindowCount=iWindowCount+1				
			Set objWindows = Nothing
			Set objCATIA=Nothing
		End If
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,sCaseName,sMenuLabel
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CATIASubMenuOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction		
		'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		If sMenuLabelTemp="CATIA V5:Export as reference" or sMenuLabelTemp="CATIA V5:Load selected level in CATIA V5" Then
			Wait 1
		ElseIf sMenuLabelTemp="CATIA V5:Load in CATIA V5" Then		
			'Creating object of [ Information ] dialog
			Set objProcessDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_LoadProcess","")
			bFlag=False
			For iCounter=1 to 180
				If Fn_UI_Object_Operations("RAC_Common_CATIASubMenuOperations","Exist",objProcessDialog,"","","") then
					Wait 1
				Else
					bFlag=True
					wait 1
					Exit For
				End If
			Next
			Set objProcessDialog=Nothing
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Load in CATIA for selected structure","","","","","")
				Call Fn_ExitTest()
			End IF
		ElseIf sMenuLabelTemp="CATIA V5:Create export spreadsheet" Then
			
			'Creating object of [ Information ] dialog
			Set objProcessDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CreateExportSpreadsheet","")
			bFlag=False
			For iCounter=1 to 120
				If Fn_UI_Object_Operations("RAC_Common_CATIASubMenuOperations","Exist",objProcessDialog,"","","") then
					Wait 1
					Call Fn_RAC_ReadyStatusSync(1)
				Else
					bFlag=True
					Call Fn_RAC_ReadyStatusSync(1)
					Exit For
				End If
			Next
			Set objProcessDialog=Nothing
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Create export spreadsheet for selected structure","","","","","")
				Call Fn_ExitTest()
			End IF
		ElseIf sMenuLabelTemp="CATIA V5:Export" Then
			
			'Creating object of [ Information ] dialog
			Set objProcessDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_ExportProcess","")
			bFlag=False
			For iCounter=1 to 240
				If Fn_UI_Object_Operations("RAC_Common_CATIASubMenuOperations","Exist",objProcessDialog,"","","") then
					Wait 1
					Call Fn_RAC_ReadyStatusSync(1)
				Else
					bFlag=True
					Call Fn_RAC_ReadyStatusSync(1)
					Exit For
				End If
			Next
			Set objProcessDialog=Nothing
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to export selected structure","","","","","")
				Call Fn_ExitTest()
			End IF
		ElseIf sMenuLabelTemp="CATIA V5:Update assembly mass properties" Then
			'Creating object of [ Information ] dialog
			Set objInformation=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_Information","")
			'Checking existance of [ Information] dialog
			If Fn_UI_Object_Operations("RAC_Common_CATIASubMenuOperations","Exist",objInformation,"","","") then
				If Fn_UI_JavaButton_Operations("RAC_Common_CATIASubMenuOperations","Click",objInformation,"jbtn_OK")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to update assembly mass properties as fail to click on [ OK ] button of [ Information ] dialog","","","","","")
					Call Fn_ExitTest()
				End If
			End IF			
			'Releasing object of [ Information] dialog
			Set objInformation=Nothing
			Call Fn_RAC_ReadyStatusSync(MAX_SYNC_ITERATIONS)
		Else
			Call Fn_RAC_ReadyStatusSync(MAX_SYNC_ITERATIONS)
		End If
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "SelectWithoutSync"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,sCaseName,sMenuLabel		
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "VeirfyExist"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,sCaseName,sMenuLabel
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "ActivateTeamcenter"
		Call Fn_RAC_SetVisible()
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "ReturnToCATIA"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,sCaseName,"CATIAReturnToCATIA"
		'Call Fn_RAC_ReadyStatusSync(1)
		Call Fn_CATIA_ReadyStatusSync(1)
End Select


