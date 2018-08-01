'! @Name 			RAC_Common_PropertiesTreeOperations 
'! @Details 		This action word is used to perform operations in property panel tree
'! @InputParam1		sAction 		: String to indicate type of action to be performed. e.g. Verify
'! @InputParam2 	sPropertyName 	:  Name of the property to be verified
'! @InputParam3		sPropertyValue 	:  Value of the property
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			13 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_PropertiesTreeOperations","RAC_Common_PropertiesTreeOperations",OneIteration,"Verify","Properties~Type","PDF"

Option Explicit
Err.Clear

'Declaring varaibles
Dim sAction,sPropertyName,sPropertyValue
Dim aPropertyName,aPropertyValue
Dim objPropertiesTree
Dim iCount,iCounter
Dim sTempValue
Dim bFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sPropertyName = Parameter("sPropertyName")
sPropertyValue = Parameter("sPropertyValue")

'Setting [ General -> Properties ] view
LoadAndRunAction "RAC_Common\RAC_Common_SetView","RAC_Common_SetView",OneIteration,"Menu","General~Properties"

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_PropertiesTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Click on [ Show Advanced Properties ] Toolbar button to display all properties of the Object
LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","ShowAdvancedProperties","",""

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_PropertiesTreeOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Creating object of [ Properties ] Tree
Set objPropertiesTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jtree_PropertiesTree","")

'Checking existance of [ Properties ] Tree
If Fn_UI_Object_Operations("RAC_Common_CreateChange","Exist", objPropertiesTree, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Properties ] tree as tree does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Spliting All properties and values
aPropertyName =Split(sPropertyName,"^",-1,1)
aPropertyValue=Split(sPropertyValue,"^",-1,1)

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "Verify"
		For iCounter=0 to ubound(aPropertyName)
			bFlag=False
			For iCount=0 to objPropertiesTree.GetROProperty("count_all_items")-1
				'Matching property with current property
				If aPropertyName(iCounter)=objPropertiesTree.GetItem(iCount) then	
					'Retrive value for specific property
					sTempValue=Cstr(objPropertiesTree.GetColumnValue(aPropertyName(iCounter),"Value"))				
					If aPropertyValue(iCounter)="{BLANK}" Then
						aPropertyValue(iCounter)=""
					ElseIf IsNumeric(aPropertyValue(iCounter)) then								
						aPropertyValue(iCounter) = Cstr(Cdbl(aPropertyValue(iCounter)))
						sTempValue= Cstr(Cdbl(sTempValue))
					End If
					'Matching Expected value with current value
					If Trim(sTempValue)=Trim(aPropertyValue(iCounter)) Then
						bFlag=True
					ElseIf IsDate(aPropertyValue(iCounter)) Then
						If Instr(1,sTempValue,aPropertyValue(iCounter)) Then
							bFlag=True
						End If
					End If
					Exit for
				End if
			Next
			If bFlag=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aPropertyName(iCounter)) & " ] property does not contain value [ " & Cstr(aPropertyValue(iCounter)) & " ]","","","","","")
				Call Fn_ExitTest()
			Else	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aPropertyName(iCounter)) & " ] property contain value [ " & Cstr(aPropertyValue(iCounter)) & " ]","","","","DONOTSYNC","")
			End If
		Next
End Select

'Closing Properties tab
LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations", oneIteration,"Close","Properties",""
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Releasing object of [ Properties ] Tree
Set objPropertiesTree=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objPropertiesTree=Nothing
 	ExitTest
End Function


