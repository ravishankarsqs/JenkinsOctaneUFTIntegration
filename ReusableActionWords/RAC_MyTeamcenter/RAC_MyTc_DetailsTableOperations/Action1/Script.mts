'! @Name 			RAC_MyTc_DetailsTableOperations
'! @Details 		Action word to perform operations on Details tab table
'! @InputParam1 	sAction 				: String to indicate what action is to be perform
'! @InputParam2 	sObject					: Object name\Column Primary value
'! @InputParam3 	sColumnName				: Column name
'! @InputParam4 	sValue					: Expected value
'! @InputParam5 	sPopupMenu				: Popup menu
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Mar 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_DetailsTableOperations","RAC_MyTc_DetailsTableOperations",OneIteration,"VerifyExist","000086/A","","",""

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sObject,sColumnName,sValue,sPopupMenu
Dim iCounter,bFlag
Dim objDetailsTable
Dim aObject
Dim sSecondaryValue
Dim bSecondaryValueFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction=Parameter("sAction")
sObject=Parameter("sObject")
sColumnName=Parameter("sColumnName")
sValue=Parameter("sValue")
sPopupMenu=Parameter("sPopupMenu")

'Creating object of Details table
Set objDetailsTable=Fn_FSOUtil_XMLFileOperations("getobject","RAC_MyTeamcenter_OR","jtbl_Details","")

'Selecting Details tab
LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select","Details",""

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_MyTc_DetailsTableOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of Details table
If Not Fn_UI_Object_Operations("RAC_MyTc_DetailsTableOperations","Exist", objDetailsTable, GBL_DEFAULT_TIMEOUT,"","") Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Details ] table as Deatils table does not exist","","","","","")
	Call Fn_ExitTest()
End If

Select Case sAction
	' - - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -- - - - - - - - - - -
	Case "VerifyExist"
		'Secondary value should be type
		aObject=Split(sObject,"^")
		If Ubound(aObject)=1 Then
			sSecondaryValue=aObject(1)
			sObject=aObject(0)
		Else	
			sSecondaryValue=""
		End If
		
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_MyTc_DetailsTableOperations","GetRowCount",objDetailsTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objDetailsTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				If sSecondaryValue<>"" Then
					If sSecondaryValue=trim(objDetailsTable.Object.getItem(iCounter).getData().getComponent().getProperty("object_type")) Then
						bSecondaryValueFlag=True
					Else
						bSecondaryValueFlag=False
					End If
				Else
					bSecondaryValueFlag=True
				End If
				
				If bSecondaryValueFlag=True Then
					If sColumnName<>"" Then
						bFlag=False
						sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
						If sColumnName="Relation" Then
							Select Case sValue
								Case "Specifications"
									sValue="IMAN_specification"
							End Select
							If trim(sValue)=trim(objDetailsTable.Object.getItem(iCounter).getData().getContext().toString()) Then
								bFlag=True
							End If
						Else
							If trim(sValue)=trim(objDetailsTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
								bFlag=True
							End If
						End If
					Else
						bFlag=True
					End If
					Exit For
				End If
			End If
		Next
		If bFlag=True Then
			IF sValue<>"" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sObject) & " ] exist under column [ Object ]","","","","","") 
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
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_MyTc_DetailsTableOperations","GetRowCount",objDetailsTable,"","","","","",""))-1		
			bFlag=False
			If Trim(objDetailsTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then				
				If sColumnName<>"" Then
					bFlag=False
					sColumnName=Fn_RAC_GetRealPropertyName(sColumnName)
					If trim(sValue)=trim(objDetailsTable.Object.getItem(iCounter).getData().getComponent().getProperty(sColumnName)) Then
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
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sValue) & " ] does not exist under column [ " & Cstr(sColumnName) & " ] against value [ " & Cstr(sObject) & " ]","","","","","") 
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified value [ " & Cstr(sObject) & " ] does not exist under column [ Object ]","","","","","") 
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
		For iCounter=0 to Cint(Fn_UI_JavaTable_Operations("RAC_MyTc_DetailsTableOperations","GetRowCount",objDetailsTable,"","","","","",""))-1
			bFlag=False
			If Trim(objDetailsTable.Object.getItem(iCounter).getData().toString())=trim(sObject) Then
				objDetailsTable.SelectCell iCounter,"Object"
				bFlag=True				
				Exit For
			End If
		Next
		
		If bFlag=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select row with value [ " & Cstr(sValue) & " ] from Details table as value [ " & Cstr(sValue) & " ] does not exist in table","","","","","")
			Call Fn_ExitTest()
		End If		
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected object [ " & Cstr(sObject) & " ] from Details table","","","",GBL_MIN_SYNC_ITERATIONS,"")
End Select

'Releasing object of [ Details ] table
Set objDetailsTable=Nothing

Function Fn_ExitTest()
	Set objDetailsTable=Nothing
	ExitTest
End Function
