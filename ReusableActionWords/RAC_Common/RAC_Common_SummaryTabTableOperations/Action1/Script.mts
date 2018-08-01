'! @Name 			RAC_Common_SummaryTabTableOperations
'! @Details 		Action word to perBusinessObject operations on New Business Object creation dialog
'! @InputParam1 	sAction 				: String to indicate what action is to be perBusinessObjected
'! @InputParam2 	sObject					: Object name\Column Primary value
'! @InputParam3 	sColumnName				: Column name
'! @InputParam4 	sValue					: Expected value
'! @InputParam5 	sPopupMenu				: Popup menu
'! @InputParam6 	sTableName				: Table name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			16 Dec 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SummaryTabTableOperations","RAC_Common_SummaryTabTableOperations",OneIteration,"VerifyExist","000086/A","","","","Customer Info"
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SummaryTabTableOperations","RAC_Common_SummaryTabTableOperations",OneIteration,"OpenWithMenuAndVerifyWindowDislpayed","AUT_5245","","","","Customer Info"	

Option Explicit
Err.Clear

'Declaring varaibles
Dim iCounter,bFlag,iOccourance,iInstance
Dim objSummaryTabTable,objDefaultWindow,objContextMenu, objWarningDialog
Dim sAction,sTableName,sObject,sColumnName,sValue,sPopupMenu
Dim sInnerTabName
Dim aPopupMenu,aObject,aValue
Dim sTemValue,sPrimaryColumnName
Dim iCount

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction=Parameter("sAction")
sObject=Parameter("sObject")
sColumnName=Parameter("sColumnName")
sValue=Parameter("sValue")
sPopupMenu=Parameter("sPopupMenu")
sTableName=Parameter("sTableName")

If sPopupMenu<>"" Then
	'Storing menu label
	sPopupMenu=Fn_RAC_GetXMLNodeValue("RAC_Common_SummaryTabTableOperations","",sPopupMenu)
   
	If sPopupMenu=False Then
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Failed to fetch value of popup menu label [ " & Cstr(Parameter("sPopupMenu")) & " ] from XML while performing menu operation","","","","DONOTSYNC","")
		Call Fn_ExitTest()
	End If
End If

'Creating object of Teamcenter Default Window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_DefaultWindow","")
'Select RMB menu
Set objContextMenu=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","wmnu_ContextMenu","")

Select Case sTableName
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "Workflow Audit Logs"
		'Selecting inner tab tab
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Workflow Information"
		Call Fn_UI_Object_Operations("RAC_Common_SummaryTabTableOperations","SetTOProperty", objDefaultWindow.JavaStaticText("jstx_SummaryTabTableHeader"),"","label",sTableName)
		'Creating Object of summary tab table
		Set objSummaryTabTable=objDefaultWindow.JavaTable("jtbl_SummaryTabTable")	
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "Reference Items"
		'Selecting inner tab tab
		LoadAndRunAction "RAC_Common\RAC_Common_InnerTabOperations","RAC_Common_InnerTabOperations",OneIteration,"Activate","Summary","Reference Data"
		Call Fn_UI_Object_Operations("RAC_Common_SummaryTabTableOperations","SetTOProperty", objDefaultWindow.JavaStaticText("jstx_SummaryTabTableHeader"),"","label",sTableName)
		'Creating Object of summary tab table
		Set objSummaryTabTable=objDefaultWindow.JavaTable("jtbl_SummaryTabTable")
End Select

'Checking existance of table
If Not Fn_UI_Object_Operations("RAC_Common_SummaryTabTableOperations","Exist", objSummaryTabTable, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ " & Cstr(sTableName) & " ] table as [ " & Cstr(sTableName) & " ] does not exist","","","","","")
	Call Fn_ExitTest()
End If


Select Case sAction
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyExist"
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then				
				If sColumnName<>"" Then
					bFlag=False
					sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
					If trim(sValue)=trim(objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
						bFlag=True
					End If
				Else
					bFlag=True
				End If
				Exit For
			End If
		Next
		If bFlag=True Then
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sObject) & " ] exist under column [ Object ]","","","","DONOTSYNC","") 
			End If	
		Else
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ Object ]","","","","","") 
			End IF	
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyNonExist"
		bFlag=False
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then				
				If sColumnName<>"" Then
					bFlag=False
					sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
					If trim(sValue)=trim(objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
						bFlag=True
					End If
				Else
					bFlag=True
				End If
				Exit For
			End If
		Next
		
		If bFlag=False Then
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sObject) & " ] does not exist under column [ Object ]","","","","DONOTSYNC","") 
			End If	
		Else
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] exist under column [ Object ]","","","","","") 
			End IF	
			Call Fn_ExitTest()			
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "Select"
		If sColumnName="" Then
			sColumnName="Customer Part Name"
		End If
		
		sPrimaryColumnName=""
		If Instr(1,sObject,"<<>>") Then
			aObject=Split(sObject,"<<>>")
			sObject=aObject(0)
			sPrimaryColumnName=aObject(1)
		End If
		
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If sPrimaryColumnName="" Then
				sTemValue=objSummaryTabTable.Object.getItem(iCounter).getData().toString()
			Else
				sPrimaryColumnName=Fn_RAC_GetRealPropertyName(sPrimaryColumnName)
				sTemValue=objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sPrimaryColumnName)
			End If
			
			If Trim(sTemValue)=trim(sObject) Then
				objSummaryTabTable.SelectCell iCounter,sColumnName
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "Cut"
		If sColumnName="" Then
			sColumnName="Customer Part Name"
		End If
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				objSummaryTabTable.SelectCell iCounter,sColumnName
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to cut row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on Add New button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objDefaultWindow,"jbtn_Cut") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to cut row with value [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ] as fail to click on [ Cut ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully cut object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "Paste"
		'Click on Paste button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objDefaultWindow,"jbtn_Paste") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to paste data in table [ " & Cstr(sTableName) & " ] as fail to click on [ Paste ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully paste copied object in table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "PopupMenuSelect"
		If sColumnName="" Then
			sColumnName="Customer Part Name"
		End If
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				objSummaryTabTable.ClickCell iCounter, sColumnName, "RIGHT"
				wait 2
				aPopupMenu = Split(sPopupMenu, ":",-1,1)
				Select Case Ubound(aPopupMenu)
						Case "0"
							 sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0))
							 objContextMenu.Select sPopupMenu
						Case "1"
							sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1))
							objContextMenu.Select sPopupMenu
						Case "2"
							sPopupMenu = objContextMenu.BuildMenuPath(aPopupMenu(0),aPopupMenu(1),aPopupMenu(2))
							objContextMenu.Select sPopupMenu
						Case Else
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to build menu path of context menu of node [" & Cstr(sObject) & "]","","","","","")
							Call Fn_ExitTest()
				End Select
							'JavaWindow("jwnd_DefaultWindow").WinMenu("wmnu_ContextMenu").Select sPopupMenu
				'wait 2
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")

    ' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "Copy"
		If sColumnName="" Then
			sColumnName="Customer Part Name"
		End If
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				objSummaryTabTable.SelectCell iCounter,sColumnName
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to cut row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on Copy button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objDefaultWindow,"jbtn_Copy") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to cut row with value [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ] as fail to click on [ Copy ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully copy object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyTableIsEmpty"
		If Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))=0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sTableName) & " ] is empty","","","","DONOTSYNC","") 	
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sTableName) & " ] is not empty\it contains some values","","","","","") 
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyTableIsNotEmpty"
		If Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))=0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sTableName) & " ] is empty it contains no values","","","","","") 
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sTableName) & " ] is not empty","","","","DONOTSYNC","") 	
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "DoubleClickCell"
		If sColumnName="" Then
			sColumnName="Customer Part Name"
		End If
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				objSummaryTabTable.DoubleClickCell iCounter,sColumnName
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to open row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully open object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyColumnExist"
		sTemValue=objSummaryTabTable.GetROProperty("columns_names")
		sTemValue=Split(sTemValue,";")
		sColumnName=SPlit(sColumnName,"~")
		For iCount=0 to Ubound(sColumnName)
			bFlag=False
			For iCounter=0 to Ubound(sTemValue)
				If sColumnName(iCount)=sTemValue(iCounter) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sColumnName(iCount)) & " ] column available in [ " & Cstr(sTableName) & " ] table","","","","DONOTSYNC","") 
					bFlag=True
					Exit For
				End If
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sColumnName(iCount)) & " ] column not available in [ " & Cstr(sTableName) & " ] table","","","","","") 
				Call Fn_ExitTest()
			End IF
		Next
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyValueExist"
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) or Trim(objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty("object_name"))=trim(sObject) Then	
				If sColumnName<>"" Then
					bFlag=False
					sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
					If trim(sValue)=trim(objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
						bFlag=True
						Exit For
					End If
				End If				
			End If
		Next
		If bFlag=True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","DONOTSYNC","") 			
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","","") 	
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyExistExt"
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
			If trim(sValue)=trim(objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
				bFlag=True
				Exit For
			End If
		Next
		If bFlag=True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","DONOTSYNC","") 
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","","") 
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "OpenWithMenuAndVerifyWindowDislpayed"
		If sColumnName="" Then
			sColumnName="Customer Part Name"
		End If

		'Select the required cell
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objSummaryTabTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				objSummaryTabTable.SelectCell iCounter,sColumnName
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform open with menu as failed to find value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Select menu option File -> Open
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileOpen"
		
		'Verify existence of dialog
		Set objWarningDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_Warning","")
		If Fn_UI_Object_Operations("RAC_Common_SummaryTabTableOperations","settoexistcheck", objWarningDialog,"", "title", sObject) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified window with title [ " & Cstr(sObject) & " ] is displayed.","","","","DONOTSYNC","") 			
			objWarningDialog.Close
			Call Fn_UI_Object_Operations("RAC_Common_SummaryTabTableOperations","settoproperty", objWarningDialog,"", "title", "Warning")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail window with title [ " & Cstr(sObject) & " ] does not exist.","","","","","") 	
			Call Fn_ExitTest()
		End If
		Set objWarningDialog = Nothing
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyColumnValueExist"
		bFlag=False
		sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			If trim(sValue)=trim(objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
				bFlag=True
				Exit For
			End If
		Next
		
		If bFlag=True Then
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ]","","","","DONOTSYNC","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sObject) & " ] exist under column [ Object ]","","","","DONOTSYNC","") 
			End If	
		Else
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ]","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as value [ " & Cstr(sValue) & " ] does not exist under column [ Object ]","","","","","") 
			End IF	
			Call Fn_ExitTest()
		End If
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "OpenInMyTeamcenter"	
		sPrimaryColumnName=""
		If Instr(1,sObject,"<<>>") Then
			aObject=Split(sObject,"<<>>")
			sObject=aObject(0)
			sPrimaryColumnName=aObject(1)
		End If
		
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If sPrimaryColumnName="" Then
				sTemValue=objSummaryTabTable.Object.getItem(iCounter).getData().toString()
			Else
				sPrimaryColumnName=Fn_RAC_GetRealPropertyName(sPrimaryColumnName)
				sTemValue=objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sPrimaryColumnName)
			End If
			
			If Trim(sTemValue)=trim(sObject) Then
				objSummaryTabTable.SelectCell iCounter,sColumnName
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to [ Open In My Teamcenter ] row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on Copy button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateBusinessObject", "Click", objDefaultWindow,"jbtn_OpenInMyTeamcenter") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to [ Open In My Teamcenter ] row with value [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ] as fail to click on [ [ Open In My Teamcenter ] ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully [ Open In My Teamcenter ] object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")	
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "SelectAll"
		If sColumnName="" Then
			sColumnName="Object"
		End If
		Dim WshShell
		bFlag=False
		sTemValue=objSummaryTabTable.GetROProperty("rows")
		If sTemValue>0 Then
			sTemValue=sTemValue-1
			objSummaryTabTable.SelectCell 0,sColumnName
			wait 1
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.SendKeys "^(a)"
			Set WshShell = Nothing
			bFlag=True
		End If
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select all rows from table [ " & Cstr(sTableName) & " ] as there are no values available in table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected all rows from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -	
	Case "SelectExt"
		If sColumnName="" Then
			sColumnName="Object"
		End If
		
		sPrimaryColumnName=""
		If Instr(1,sObject,"<<>>") Then
			aObject=Split(sObject,"<<>>")
			sObject=aObject(0)
			sPrimaryColumnName=aObject(1)
		End If
		
		If Instr(1,sObject,"@") Then
			aObject=Split(sObject,"@")
			sObject=aObject(0)
			iInstance=aObject(1)
		Else
			iInstance=1
		End If
		
		iOccourance=1
		
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_Common_SummaryTabTableOperations","GetRowCount",objSummaryTabTable,"","","","","",""))-1		
			bFlag=False
			If sPrimaryColumnName="" Then
				sTemValue=objSummaryTabTable.Object.getItem(iCounter).getData().toString()
			Else
				sPrimaryColumnName=Fn_RAC_GetRealPropertyName(sPrimaryColumnName)
				sTemValue=objSummaryTabTable.Object.getItem(iCounter).getData().getComponent().getProperty(sPrimaryColumnName)
			End If
			
			If Trim(sTemValue)=trim(sObject) Then				
				If Cint(iInstance)=Cint(iOccourance) Then
					objSummaryTabTable.SelectCell iCounter,sColumnName
					bFlag=True				
					Exit For
				End If
				iOccourance=iOccourance+1
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select row with value [ " & Cstr(sValue) & " ] from table [ " & Cstr(sTableName) & " ] as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected object [ " & Cstr(sObject) & " ] from table [ " & Cstr(sTableName) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
End Select

'Releasing object of [ Summary tab ] table
Set objSummaryTabTable=Nothing
Set objDefaultWindow =Nothing

Function Fn_ExitTest()
	Set objSummaryTabTable=Nothing
	Set objDefaultWindow =Nothing
	ExitTest
End Function

