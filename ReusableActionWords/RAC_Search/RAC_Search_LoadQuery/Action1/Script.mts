'! @Name 			RAC_Search_LoadQuery
'! @Details 		This actionword is used to load search query
'! @InputParam1		sQueryPath : Search query path
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			10 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Search\RAC_Search_LoadQuery","RAC_Search_LoadQuery",OneIteration,"System Defined Searches~Item"

Option Explicit
Err.Clear

'Declaring variables
Dim objChangeSearch
Dim aQueryPath
Dim sQueryPath

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sQueryPath = Parameter("sQueryPath")

aQueryPath = Split(sQueryPath, "~", -1, 1)

'Setting search view
LoadAndRunAction "RAC_Common\RAC_Common_SetView","RAC_Common_SetView",OneIteration,"Menu","Teamcenter~Search"

'Click on [Select a Search] toolbar button under Search Criteria Panel
LoadAndRunAction "RAC_Common\RAC_Common_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","SelectASearch","",""

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Search_LoadQuery"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Creating object of change search dialog
Set objChangeSearch=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Search_OR","jwnd_ChangeSearch","")

'Capture business functionality start time	
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Load Search Query","","Query Name",sQueryPath)	

'Checked Existance of Search Result Tree
If Fn_UI_Object_Operations("RAC_Search_LoadQuery","Exist",objChangeSearch,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to load search query [ " & Cstr(sQueryPath) & " ] as [ Change Search ] dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Checked Existance of Search Result Tree
If Fn_UI_Object_Operations("RAC_Search_LoadQuery","Exist",objChangeSearch.JavaTree("jtree_SearchOptions"),"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to load search query [ " & Cstr(sQueryPath) & " ] as [ Search Option ] Tree does not exist","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Expand parent node of search query
If Fn_UI_JavaTree_Operations("RAC_Search_LoadQuery","Expand",objChangeSearch,"jtree_SearchOptions",aQueryPath(0),"","")=False THen
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to load search query [ " & Cstr(sQueryPath) & " ] as to expand [ " & Cstr(aQueryPath(0)) & " ] node from search options tree","","","","","")
	Call Fn_ExitTest()
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
wait 1
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Selecting search query
If Fn_UI_JavaTree_Operations("RAC_Search_LoadQuery","Select",objChangeSearch,"jtree_SearchOptions",sQueryPath,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to load search query [ " & Cstr(sQueryPath) & " ] as to select query\node from search options tree","","","","","")
	Call Fn_ExitTest()
End IF
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)

'Click on [ OK ] button
If Fn_UI_JavaButton_Operations("RAC_Search_LoadQuery", "Click", objChangeSearch, "jbtn_OK") =True Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected and load search query [ " & Cstr(sQueryPath) & " ]","","","",GBL_MIN_SYNC_ITERATIONS,"")
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Load Search Query","","Query Name",sQueryPath)	
Else
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select and load search query [ " & Cstr(sQueryPath) & " ] as fail to click on [ OK ] button of load query\change search dialog","","","","","")
	Call Fn_ExitTest()
End If

If Err.Number <> 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select\laod search query [ " & Cstr(sQueryPath) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Releasing [ Change Search ] dialog object			
Set objChangeSearch  = Nothing

Function Fn_ExitTest()
	'Releasing [ Change Search ] dialog object			
	Set objChangeSearch  = Nothing
	ExitTest
End Function

