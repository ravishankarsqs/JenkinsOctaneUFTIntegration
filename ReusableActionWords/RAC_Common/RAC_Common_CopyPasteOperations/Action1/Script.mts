'! @Name 			RAC_Common_CopyPasteOperations
'! @Details 		Action word to perform copy paste operations
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sNodeToCopy 		: Name of the node to be copied
'! @InputParam3 	sNodeCopyFrom 		: Node to be copied from
'! @InputParam4 	sNodeToPaste		: Name of the node to be pasted in
'! @InputParam5 	sPasteAsRelation	: Paste as relation
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			09 Jan 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CopyPasteOperations","RAC_Common_CopyPasteOperations", oneIteration, "CopyPaste",GBL_TEST_CASE_FOLDER_COMPLETE_NAME &"~"& Datatable.Value("PartRevision"),"NavigationTree",GBL_TEST_CASE_FOLDER_COMPLETE_NAME &"~"& Datatable.Value("DesignRevision") &"~Represented By",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction, sNodeToCopy, sNodeCopyFrom,sNodeToPaste, sPasteAsRelation

GBL_CURRENT_EXECUTABLE_APP="RAC"

'Set parameter values in local variables
sAction= Parameter("sAction")
sNodeToCopy = Parameter("sNodeToCopy")
sNodeCopyFrom = Parameter("sNodeCopyFrom")
sNodeToPaste = Parameter("sNodeToPaste")
sPasteAsRelation = Parameter("sPasteAsRelation")

If sNodeToCopy <> "" Then
	Select Case sNodeCopyFrom
		Case "NavigationTree",""
			If sAction="CopyPaste" or sAction="CopyPasteWithRelation" Then
				'Select the node to be copied
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "ExpandAndSelect",sNodeToCopy,""
			ElseIf sAction="CopyPasteExt" Then
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select",sNodeToCopy,""
			ElseIf sAction="BasicCopyPaste" Then
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select",sNodeToCopy,""
			End If
	End Select
End If

Select Case sAction
	Case "BasicCopyPaste"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditCopy"
		Select Case sNodeCopyFrom
			Case "NavigationTree",""
				'Select the node to be copied
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select",sNodeToPaste,""
		End Select
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditPaste"
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully pasted copied data on [ " & Cstr(sNodeToPaste) & " ] node","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		
	Case "CopyPaste","CopyPasteExt"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditCopy"
		Select Case sNodeCopyFrom
			Case "NavigationTree",""
				'Select the node to be copied
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "ExpandAndSelect",sNodeToPaste,""
		End Select
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditPaste"
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully pasted copied data on [ " & Cstr(sNodeToPaste) & " ] node","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		
	Case "Paste"
		Select Case sNodeCopyFrom
			Case "NavigationTree",""
				'Select the node to be copied
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "ExpandAndSelect",sNodeToPaste,""
		End Select
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditPaste"
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully pasted copied data on [ " & Cstr(sNodeToPaste) & " ] node","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		
	Case "PasteWithCutPasteWarning","PasteWithCopyPasteWarning"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditPaste"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CopyPasteOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction
		
		If JavaWindow("jwnd_CutPasteWarning").Exist(10) Then
			'Click on Yes button
			If Fn_UI_JavaButton_Operations("RAC_Common_CopyPasteOperations", "Click", JavaWindow("jwnd_CutPasteWarning"),"jbtn_Yes") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Yes ] button from [ Cut/Paste Warning ] dialog","","","","","")
				Call Fn_ExitTest()
			End If			
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully clicked on [ Yes ] button of [ Cut/Paste Warning ] dialog and pasted copied data on selected node","","","",GBL_MICRO_SYNC_ITERATIONS,"")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail as [ Cut/Paste Warning ] dialog does not appears after performing paste operation","","","","","")
		End If
		
	Case "CopyPasteWithRelation"
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditCopy"
		Select Case sNodeCopyFrom
			Case "NavigationTree",""
				'Select the node to be copied
				LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "ExpandAndSelect",sNodeToPaste,""
		End Select
		LoadAndRunAction "RAC_Common\RAC_Common_EditPasteSpecial","RAC_Common_EditPasteSpecial", oneIteration, "PasteSpecial","menu",sPasteAsRelation,""
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully pasted copied data on [ " & Cstr(sNodeToPaste) & " ] node","","","",GBL_MICRO_SYNC_ITERATIONS,"")
End Select

