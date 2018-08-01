'! @Name 			RAC_PSE_ObjectReplaceOperations
'! @Details 		Action word to Replace object in assembly
'! @InputParam1 	sAction 			: Action to be performed e.g. AutoReplaceBasic
'! @InputParam2 	sInvokeOption 		: Method to invoke Replace dialog e.g. menu
'! @InputParam3 	sCopyNodePath		: Table node path to copy
'! @InputParam4 	sReplaceNodePath	: Table node path to replace
'! @InputParam5 	sRevision 			: Revision ID
'! @InputParam6 	sViewType		 	: Replace view type
'! @InputParam7 	sReplaceOption	 	: Replace option
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			14 Dec 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_PSE_ObjectReplaceOperations","RAC_PSE_ObjectReplaceOperations",OneIteration,"copyandreplace","","0000313/AA-Asm Cockpit~0000316/AA-Asm Cockpit","0000313/AA-Asm Cockpit~0000317/AA-Asm Cockpit","","",""
'! @Example 		LoadAndRunAction "RAC_Common\RAC_PSE_ObjectReplaceOperations","RAC_PSE_ObjectReplaceOperations",OneIteration,"searchandreplace","menu","0000313~Asm Cockpit","0000313/AA-Asm Cockpit~0000317/AA-Asm Cockpit","AA","","The original"

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sCopyNodePath,sReplaceNodePath,sRevision,sViewType,sReplaceOption
Dim objReplace
Dim aObjectInfo

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
'In "CopyAndReplace" case use this paramenter to pass full path of node which user wants to copy
'In "FindAndReplace" case use this paramenter to pass Item ID ~ Item Name which user wants to search
sCopyNodePath = Parameter("sCopyNodePath")
sReplaceNodePath = Parameter("sReplaceNodePath")
sRevision = Parameter("sRevision")
sViewType = Parameter("sViewType")
sReplaceOption = Parameter("sReplaceOption")

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ObjectReplaceOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Replace Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to copy and  Replace obejct
	Case "copyandreplace"
		If sCopyNodePath<>"" Then
			'Selecting node from table to copy			
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",sCopyNodePath,"","",""
			'Calling menu Edit->Copy
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditCopy"
		End If
		
		If sReplaceNodePath<>"" Then
			'Selecting node from table to replace			
			LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",sReplaceNodePath,"","",""
		End If
		
		'Calling menu Edit->Replace
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditReplace"
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Replace Operations",sAction,"","")
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ObjectReplaceOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Replace selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Replaced selected object in assembly","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Search and Replace obejct
	Case "searchandreplace","searchandreplaceinsplitbom"	
		If sReplaceNodePath<>"" Then
			If Lcase(sAction)="searchandreplace" Then
				'Selecting node from table to replace			
				LoadAndRunAction "RAC_StructureManager\RAC_PSE_BOMTableOperations","RAC_PSE_BOMTableOperations", oneIteration, "Select",sReplaceNodePath,"","",""
			ElseIf Lcase(sAction)="searchandreplaceinsplitbom" Then
				sReplaceNodePath=Split(sReplaceNodePath,"<<>>")
				LoadAndRunAction "RAC_StructureManager\RAC_PSE_SplitBOMTableOperations","RAC_PSE_SplitBOMTableOperations", oneIteration, "Select",sReplaceNodePath(0),"","","",sReplaceNodePath(1)				
			End If			
		End If
		
		'inoke Replace dialog
		Select Case LCase(sInvokeOption)
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "menu"
				LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditReplace..."
			' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
			Case "nooption"
				'Use this invoke option when user wants to invoke Replace dialog from outside function
		End Select

		'Creating object of [ Replace ] dialog
		Set objReplace=Fn_FSOUtil_XMLFileOperations("getobject","RAC_StructureManager_OR","jdlg_Replace","")
		
		GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_PSE_ObjectReplaceOperations"
		GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

		'Checking existance of Replace dialog
		If Fn_UI_Object_Operations("RAC_PSE_ObjectReplaceOperations", "Exist", objReplace, GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Replace ] dialog as dialog does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'sCopyNodePath= Item ID ~ Item Name
		aObjectInfo=Split(sCopyNodePath,"~")
		'Searching replace object
		If Fn_UI_JavaEdit_Operations("RAC_PSE_ObjectReplaceOperations","Set",objReplace,"jedt_ItemID",aObjectInfo(0))=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to replace selected object as fail to search replace object of id [ " & Cstr(aObjectInfo(0)) & " ]","","","","","")	
			Call Fn_ExitTest()
		End If
		'Activating Item ID field
		Call Fn_UI_JavaEdit_Operations("RAC_PSE_ObjectReplaceOperations","activate",objReplace,"jedt_ItemID","")
		Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		
		'Selecting revision ID
		If sRevision <> "" Then
			If Fn_UI_JavaList_Operations("RAC_PSE_ObjectReplaceOperations", "Select", objReplace,"jlst_RevisionID",sRevision, "", "")=False Then		
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to replace selected object as fail to select replace object revision id value [ " & Cstr(sRevision) & " ]","","","","","")	
				Call Fn_ExitTest()
			End If		
		End If
		'Selecting View Type
		If sViewType <> "" Then
			If Fn_UI_JavaList_Operations("RAC_PSE_ObjectReplaceOperations", "Select", objReplace,"jlst_ViewType",sViewType, "", "")=False Then		
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to replace selected object as fail to select replace object veiw type value [ " & Cstr(sViewType) & " ]","","","","","")	
				Call Fn_ExitTest()
			End If		
		End If
		'Selecting replace option
		If sReplaceOption<>"" Then
			If Fn_UI_Object_Operations("RAC_PSE_ObjectReplaceOperations","settoproperty",objReplace.JavaRadioButton("jrdb_Replace"),"","attached text",sReplaceOption)=False Then	
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to replace selected object as replace option [ " & Cstr(sReplaceOption) & " ] does not exist on replace dialog","","","","","")	
				Call Fn_ExitTest()
			End IF
			If Fn_UI_JavaRadioButton_Operations("RAC_PSE_ObjectReplaceOperations", "Set", objReplace, "jrdb_Replace", "ON") =False then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to replace selected object as fail to select replace option [ " & Cstr(sReplaceOption) & " ]","","","","","")	
				Call Fn_ExitTest()
			End If
		End If
		
		If objReplace.JavaButton("jbtn_OK").Exist(3) Then					
			'Click on [ OK ] button
			If Fn_UI_JavaButton_Operations("RAC_PSE_ObjectReplaceOperations", "Click", objReplace,"jbtn_OK")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to replace object as failed to click on [ OK ] button on [ Replace ] dialog","","","","","")
				Call Fn_ExitTest()
			End If			
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Replace Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Replace selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Replaced selected object in assembly with object id [ " & Cstr(aObjectInfo(0)) & " ]","","","","","")
		End If		
End Select

'Releasing object
Set objReplace=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objReplace=Nothing
	ExitTest
End Function
