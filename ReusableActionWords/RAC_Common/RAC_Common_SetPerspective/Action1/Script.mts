'! @Name 			RAC_Common_SetPerspective
'! @Details 		To set a perspective in Teamcenter
'! @InputParam1 	sModule : Name of the perspective to be set
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			26 Mar 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SetPerspective","RAC_Common_SetPerspective",OneIteration,"My Teamcenter"

Option Explicit
Err.Clear

'Declaring variables
Dim sModuleName
Dim objOpenPerspective,objOpenPerspectiveTable,objTable
Dim iSelectionIndex,iRowCount,iCounter
Dim sCurrentModuleName
Dim bReturn

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sModuleName=Parameter("sModuleName")

'Invoking menu Window -> Open Perspective -> Other
LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","WindowOpenPerspectiveOther"

Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Set Perspective","","Module Name",sModuleName)

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_SetPerspective"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

'Creating object of [ Open Perspective ] dialog
Set objOpenPerspective=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_OpenPerspective","")
Set objOpenPerspectiveTable = objOpenPerspective.JavaTable("jtbl_Table")

'Checking existance of [ Open Perspective ] dialog
If Fn_UI_Object_Operations("RAC_Common_SetPerspective","Exist",objOpenPerspective ,"","","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set perspective [ " & Cstr(sModuleName) & " ]","","","","","")
	Call Fn_ExitTest()
End If

'Select the First Row of the Module Table
objOpenPerspectiveTable.SelectCell "#0", "#0"
'Type in the First Letter of Modult Name
objOpenPerspectiveTable.Type Left(sModuleName, 1)

iSelectionIndex = 0
Set objTable = objOpenPerspectiveTable.Object
iSelectionIndex = objTable.getFocusIndex()
Set objTable = Nothing

'find the row count
iRowCount = Fn_UI_Object_Operations("RAC_Common_SetPerspective","getroproperty",objOpenPerspective.JavaTable("jtbl_Table"),"","rows","")
         
For iCounter = iSelectionIndex to iRowCount -1	
	'Retrive Module Name data
	sCurrentModuleName = objOpenPerspectiveTable.GetCellData(cstr(iCounter),"0")
	
	If Trim(sCurrentModuleName) = Trim(sModuleName) Then
		'Selecting perspective from table
		objOpenPerspectiveTable.SelectCell cstr(iCounter),"0"
		If Err.Number <  0 Then
            Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set perspective [ " & Cstr(sModuleName) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		'Clicking on [ OK ] button
		If Fn_UI_Object_Operations("RAC_Common_SetPerspective", "Exist",objOpenPerspective.JavaButton("jbtn_OK"),"","","") then
			If Cint(Fn_UI_Object_Operations("RAC_Common_SetPerspective","getroproperty",objOpenPerspective.JavaButton("jbtn_OK"),"","enabled",""))=0 Then
				objOpenPerspectiveTable.SelectCell cInt(iCounter),0
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		Exit for		
	End If		
Next
If iCounter = iRowCount Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set perspective [ " & Cstr(sModuleName) & " ]","","","","","")
	Call Fn_ExitTest()
End If
	
If Fn_UI_Object_Operations("RAC_Common_SetPerspective", "Exist",objOpenPerspective.JavaButton("jbtn_OK"),GBL_MIN_MICRO_TIMEOUT,"","") then	
	Call Fn_UI_JavaButton_Operations("RAC_Common_SetPerspective","Click",objOpenPerspective,"jbtn_OK")
	'Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
End IF

If Lcase(sModuleName)="project" Then
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	If JavaWindow("jwnd_DefaultWindow").Dialog("text:=non-Project Administrator Access").Exist(6) Then
		JavaWindow("jwnd_DefaultWindow").Dialog("text:=non-Project Administrator Access").WinButton("text:=OK").Click
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	ElseIf Dialog("text:=non-Project Administrator Access").Exist(6) Then
		Dialog("text:=non-Project Administrator Access").WinButton("text:=OK").Click
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End IF
End If

If Lcase(sModuleName)="organization" Then
	Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	If JavaWindow("jwnd_DefaultWindow").JavaWindow("title:=non-DBA Access").Exist(6) Then
		JavaWindow("jwnd_DefaultWindow").JavaWindow("title:=non-DBA Access").JavaButton("label:=OK").Click
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	End IF
End If

If Err.Number < 0 Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set perspective [ " & Cstr(sModuleName) & " ]","","","","","")
	Call Fn_ExitTest()
Else
	Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Set Perspective","","Module Name",sModuleName)
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully set perspective [ " & Cstr(sModuleName) & " ]","","","","DONOTSYNC","")
End If
Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)


'Releasing object of [ Open Perspective ] dialog
Set objOpenPerspectiveTable=Nothing
Set objOpenPerspective=Nothing

Function Fn_ExitTest()
	'Releasing object of [ Open Perspective ] dialog
	Set objOpenPerspectiveTable=Nothing
	Set objOpenPerspective =Nothing
	ExitTest
End Function

