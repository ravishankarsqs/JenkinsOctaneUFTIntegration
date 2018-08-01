'! @Name RAC_Search_SpecifySearchDetailsAndInvoke
'! @Details To search object by using criteria
'! @InputParam1. sSearchQuery : Search query path
'! @InputParam2. sSearchCriteria : Advance Search Criteria
'! @InputParam3. sSearchCriteriaValue : Advance Search Criteria value
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 09 Dec 2015
'! @Version 1.0
'! @Example LoadAndRunAction "RAC_Search\RAC_Search_SpecifySearchDetailsAndInvoke","RAC_Search_SpecifySearchDetailsAndInvoke",OneIteration,"System Defined Searches~General...","Name~Type","AutomatedTest~Folder"

Option Explicit

'Variable Declaration
Dim bFlag
Dim sAction
Dim iCounter, iCount
Dim aSearchCriteria,aSearchCriteriaValue
Dim objDefaultWindow,objSelectType,objIntNoOfObjects
Dim sSearchQuery,sSearchCriteria,sSearchCriteriaValue,sNode
Dim bWildCardSearch
Dim aTempValue

bWildCardSearch=False

GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading input parameter values
sSearchQuery = Parameter("sSearchQuery")
sSearchCriteria = Parameter("sSearchCriteria")
sSearchCriteriaValue = Parameter("sSearchCriteriaValue")

'Creating object of [ DefaultWindow ] window
Set objDefaultWindow=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jwnd_SearchDefaultWindow","")

'Selecting Query for specific search results
If sSearchQuery<>"" Then
	LoadAndRunAction "RAC_Search\RAC_Search_LoadQuery","RAC_Search_LoadQuery",OneIteration,sSearchQuery
Else
	If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","settoexistcheck",objDefaultWindow.JavaButton("jbtn_More"),"1", "label", "More...>>>") Then
		If Fn_UI_JavaButton_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Click",objDefaultWindow, "jbtn_More")=False Then
	   		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform search operation as fail to click on [ More ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	End IF
End If

'Checking existance of [ DefaultWindow ] window
If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Exist",objDefaultWindow,"", "", "") Then

	'Clearing all search fields
	LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", "Clearallsearchfields", "RAC_Common_TLB",""
	
	GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Search_SpecifySearchDetailsAndInvoke"
	GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= Cstr(sSearchCriteria)

	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Specify Search Details And Invoke","",sSearchCriteria,sSearchCriteriaValue)
	'
	aSearchCriteria = Split(sSearchCriteria,"~")
	aSearchCriteriaValue = Split(sSearchCriteriaValue,"~")

	For iCounter=0 to Ubound(aSearchCriteria)
		sAction=aSearchCriteria(iCounter)
		If aSearchCriteriaValue(iCounter)="*" Then
			bWildCardSearch=True
		End If
		Select Case sAction
			'Edit Box
			Case "Name","ID","Item ID","Keyword","Alternate Identifier","Opportunity ID","Cusomter Part Name","Description","Customer Number","Customer Part Revision","Revision","UOM","3D Design Required"
				objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", sAction & ":"		
				If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Exist", objDefaultWindow.JavaEdit("jedt_SearchEditBox"),"","","") Then
					If Fn_UI_JavaEdit_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Set",objDefaultWindow,"jedt_SearchEditBox",aSearchCriteriaValue(iCounter))=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] value as [ " & Cstr(aSearchCriteriaValue(iCounter)) & " ]","","","","","")
						Call Fn_ExitTest()
					End If
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] as search criteria value does not exist","","","","","")
					Call Fn_ExitTest()
				End If			
			'Drop Down list
			Case "Type","Dataset Type","Document Type","Owning User","Owning Group","Intended Tooling Maturity","Change Type"
				objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", sAction & ":"
				If Fn_UI_JavaButton_Operations( "RAC_Search_SpecifySearchDetailsAndInvoke","Click",objDefaultWindow,"jbtn_SearchMultipleDropDown" )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] as fail to click on Drop down button against the Search criteria field","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				If Fn_UI_JavaEdit_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Set",objDefaultWindow,"jedt_SearchEditBox",aSearchCriteriaValue(iCounter))=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] value as [ " & Cstr(aSearchCriteriaValue(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				wait 3
				Set objSelectType=description.Create()
				objSelectType("path").Value="Tree;Composite;Composite;Shell;Shell;"
				objSelectType("developer name").RegularExpression=True
				objSelectType("developer name").Value="com.teamcenter.rac.common.lov.common.LOVRow.*"
				objSelectType("displayed").Value=1					
				Set objIntNoOfObjects = objDefaultWindow.ChildObjects(objSelectType)
				If objIntNoOfObjects.count=0 Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] value as [ " & Cstr(aSearchCriteriaValue(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
				objIntNoOfObjects(0).Activate aSearchCriteriaValue(iCounter)
				Set objSelectType=nothing
				Set objIntNoOfObjects =nothing
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

			Case "Is Migrated?"	
				objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", sAction & ":"
				If Fn_UI_JavaList_Operations("RAC_Search_SpecifySearchDetailsAndInvoke", "Exist", objDefaultWindow,"jlast_SearchList",aSearchCriteriaValue(iCounter), "", "") Then
					If Fn_UI_JavaList_Operations("RAC_Search_SpecifySearchDetailsAndInvoke", "Select", objDefaultWindow,"jlast_SearchList",aSearchCriteriaValue(iCounter), "", "")=False Then
						Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Fail to select value [ "+aSearchCriteriaValue(iCounter)+" ] from List of Values list")
						Set objDefaultWindow=Nothing
					End If
				Else
					Call Fn_WriteLogFile(Environment.Value("TestLogFile"),"FAIL : Value [ "+aSearchCriteriaValue(iCounter)+" ] not found in List of Values list")
					Set objDefaultWindow=Nothing
				End If
				
			Case "Release Status"
				objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", sAction & ":"
				If Fn_UI_JavaEdit_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Set",objDefaultWindow,"jedt_SearchEditBox",aSearchCriteriaValue(iCounter))=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] value as [ " & Cstr(aSearchCriteriaValue(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
'				objDefaultWindow.JavaWindow("Shell").JavaTree("Tree").Click 0,1,"LEFT"
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
			Case "Created After","Created Before","Created Before WithExtraSync"
				If sAction<>"Created Before WithExtraSync"Then
					objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", sAction & ":"
				Else
					objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", Replace(sAction," WithExtraSync","") & ":"
				End If
				
				If Fn_UI_JavaEdit_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","settext",objDefaultWindow,"jedt_SearchEditBox",aSearchCriteriaValue(iCounter))=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] value as [ " & Cstr(aSearchCriteriaValue(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
			'Drop Down list
			Case "Raw Material Name"
				objDefaultWindow.JavaStaticText("jstx_SearchType").SetTOProperty "label", sAction & ":"
				If Fn_UI_JavaButton_Operations( "RAC_Search_SpecifySearchDetailsAndInvoke","Click",objDefaultWindow,"jbtn_SearchMultipleDropDown" )=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] as fail to click on Drop down button against the Search criteria field","","","","","")
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
				'Verify existence of name tree
				If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","Exist", objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree"), GBL_DEFAULT_TIMEOUT,"","")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] as value selection tree does not exist","","","","","")
					Call Fn_ExitTest()
				End If
				
				aSearchCriteriaValue(iCounter)=Replace(aSearchCriteriaValue(iCounter),"^","~")
				'Expand the name tree till last parent hierarchy
				aTempValue = Split(aSearchCriteriaValue(iCounter), "~")
				For iCount = 0 To Ubound(aTempValue) - 1 Step 1
					If iCount = 0 Then
						sNode = aTempValue(iCount)
					Else
						sNode = sNode & "~" & aTempValue(iCount)
					End If
					objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree").Expand sNode
'					wait 0,500
					Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
					If Err.Number<0 then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set search criteria [ " & Cstr(sAction) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
						Call Fn_ExitTest()
					End If
				Next
				
				'Select the name node
				objDefaultWindow.JavaWindow("jwnd_LOVTreeShell").JavaTree("jtree_LOVTree").Activate aSearchCriteriaValue(iCounter)
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
			Case "VerifySearchAttributes"
				For iCount=0 to ubound(aSearchCriteriaValue)
					If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","settoproperty",objDefaultWindow.JavaStaticText("jstx_SearchType"),"","label",aSearchCriteriaValue(iCount) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aSearchCriteriaValue(iCount)) & " ] attribute does not exist\available on search query page","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke", "Exist", objDefaultWindow.JavaStaticText("jstx_SearchType"),"","","") Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aSearchCriteriaValue(iCount)) & " ] attribute available on search query page","","","","DONOTSYNC","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aSearchCriteriaValue(iCount)) & " ] attribute does not exist\available on search query page","","","","","")
						Call Fn_ExitTest()
					End If
				Next
			'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
			Case "VerifySearchAttributesNonExist"
				For iCount=0 to ubound(aSearchCriteriaValue)
					If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke","settoproperty",objDefaultWindow.JavaStaticText("jstx_SearchType"),"","label",aSearchCriteriaValue(iCount) & ":")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aSearchCriteriaValue(iCount)) & " ] attribute does not exist\available on search query page","","","","","")
						Call Fn_ExitTest()
					End IF
					
					If Fn_UI_Object_Operations("RAC_Search_SpecifySearchDetailsAndInvoke", "Exist", objDefaultWindow.JavaStaticText("jstx_SearchType"),"","","")=False Then
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aSearchCriteriaValue(iCount)) & " ] attribute does not exist\available  on search query page","","","","DONOTSYNC","")
					Else
						Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aSearchCriteriaValue(iCount)) & " ] attribute exist\available on search query page","","","","","")
						Call Fn_ExitTest()
					End If
				Next
					
		End Select
	Next
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to paerform advance search operation for search criteria [ " & Cstr(sSearchCriteria) & " ] with value [ " & Cstr(sSearchCriteriaValue) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Click on search toolbar button
If sAction = "VerifySearchAttributes" or sAction = "VerifySearchAttributesNonExist" Then
	'Nothing do anything
Else
	LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","ExecuteAndDisplaySearchResults","RAC_Common_TLB",""
	If  sAction<>"Created Before WithExtraSync" Then
 		Wait 10
 	End If
 	Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
End If

If Err.Number<0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to paerform advance search operation for search criteria [ " & Cstr(sSearchCriteria) & " ] with value [ " & Cstr(sSearchCriteriaValue) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Specify Search Details And Invoke","",sSearchCriteria,sSearchCriteriaValue)

Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully paerformed advance search operation for search criteria [ " & Cstr(sSearchCriteria) & " ] with value [ " & Cstr(sSearchCriteriaValue) & " ]","","","","","")

Set objDefaultWindow=Nothing

Function Fn_ExitTest()
	Set objDefaultWindow = Nothing 
	ExitTest
End Function

