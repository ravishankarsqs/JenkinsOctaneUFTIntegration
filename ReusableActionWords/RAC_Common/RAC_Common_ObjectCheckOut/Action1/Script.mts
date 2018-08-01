'! @Name 			RAC_Common_ObjectCheckOut
'! @Details 		This action word is used to perform check out operation on Objects
'! @InputParam1 	sAction 		: Action to be performed e.g. CheckOut
'! @InputParam2 	sInvokeOption 	: Method to invoke checkout e.g. menu
'! @InputParam3 	sPerspective 	: Perspective name
'! @InputParam4 	sChangeID 		: Change id if to be provided
'! @InputParam5 	sComment 		: Comments for the change
'! @InputParam6 	sExport 		: Export checkbox option
'! @InputParam7 	sOverwrite 		: Overwrite checkbox option
'! @InputParam8 	sButton 		: Button name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			23 Jun 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectCheckOut","RAC_Common_ObjectCheckOut",OneIteration,"CheckOut", "menu", "myteamcenter", "", "Checkout","ON","ON",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sInvokeOption,sPerspective,sChangeID,sComment,sExport,sOverwrite,sObjectName,sButton
Dim objCheckOut,objCheckingOut
Dim bFlag 

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action input parameters in local variables
sAction = Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sChangeID = Parameter("sChangeID")
sComment = Parameter("sComment")
sExport = Parameter("sExport")
sOverwrite = Parameter("sOverwrite")
sButton = Parameter("sButton")

'Creating Object of [ Checking Out ] dialog
Set objCheckingOut=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CheckingOut","")

If sPerspective="" Then
	'Get active perspective name
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

'Creating Object of [ Check out ] dialog
Select Case Lcase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","systemsengineering",""
		Set objCheckOut=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_CheckOut","")
End Select

Select Case lCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Invoke Check Out dialog From Menu
	Case "menu"
		If sAction="CancelCheckOut" Then
			'Invoke menu
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration, "Select", "ToolsCancelCheckOut"
		Else
			'Invoke menu
			LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations", oneIteration, "Select", "ToolsCheckOut"
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Invoke Check out dialog From Summary tab
	Case "summarytabtoolbar"
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select", "Summary", ""
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		LoadAndRunAction "RAC_Common\RAC_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click", "CheckOut", "",""
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to Invoke Check out dialog From Summary tab
	Case "viewertabtoolbar"
		LoadAndRunAction "RAC_Common\RAC_Common_TabFolderWidgetOperations","RAC_Common_TabFolderWidgetOperations",oneIteration,"Select", "Viewer", ""
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		LoadAndRunAction "RAC_Common\RAC_ToolbarOperations","RAC_Common_ToolbarOperations",OneIteration,"Click","CheckOut", "",""
End Select

Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectCheckOut"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

If sAction="CheckOutForm" Then
	'Checking existance of [ Check Out ] dialog
	If  Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist", objCheckOut,GBL_MIN_TIMEOUT,"","")=False Then
		'Checking existance of [ Check Out ] dialog
		If Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist", objCheckingOut,GBL_MIN_MICRO_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & " ] operation as [ Check Out ] dialog does not exist","","","","","")
			Call Fn_ExitTest()
		Else
			Set objCheckOut = objCheckingOut
			bFlag= True
		End If
	End If
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Check Out",sAction,"","")

Select Case sAction
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to perform basic check out of object
	Case "CheckOut","CheckOutWithError"
'		'Setting Change ID
'		If sChangeID<>"" Then
'			If Fn_UI_JavaEdit_Operations("RAC_Common_ObjectCheckOut", "Set",  objCheckOut, "jedt_ChangeID", sChangeID )=False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to enter value in [ Change ID ] edit box","","","","","")
'				Call Fn_ExitTest()
'			End If
'		End If
'		'Setting Comments
'		If sComment<>"" Then
'			If Fn_UI_JavaEdit_Operations("RAC_Common_ObjectCheckOut", "Set",  objCheckOut, "jedt_Comments", sComment )=False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to enter value in Comments edit box","","","","","")
'				Call Fn_ExitTest()
'			End If
'		End If
'		'Setting Export Option
'		If sExport<>"" Then
'			If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectCheckOut", "Set", objCheckOut, "jckb_ExportDtsetOnCheckOut", sExport)=False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to select checkbox for Export Dataset","","","","","")
'				Call Fn_ExitTest()
'			End If
'			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
'		End If
'		'Setting Overwrite option
'		If sOverwrite<>"" Then
'			If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectCheckOut", "Set", objCheckOut, "jckb_OverwriteExistingFiles", sOverwrite)=False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to select checkbox for Overwrite existing files","","","","","")
'				Call Fn_ExitTest()
'			End If
'			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
'		End If
'
'		'Click on Yes button to Check out Object
'		If bFlag= True Then
'			If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_OK") = False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to click on OK button","","","","","")
'				Call Fn_ExitTest()
'			End If
'		Else
'			If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_Yes") = False Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to click on Yes button","","","","","")
'				Call Fn_ExitTest()
'			End If
'		End If
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Check Out selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		End If		
		
		If sAction="CheckOutWithError" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully click on [ Yes ] button of Check Out dialog","","","","","")			
		Else
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
'			'Checking existance of [ Check Out ] dialog
'			If Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist",objCheckOut,GBL_MICRO_TIMEOUT,"","")=True Then
'				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Check Out selected object as error apeared after check out operation","","","","","")
'				Call Fn_ExitTest()
'			End If
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Check Out selected object","","","","DONOTSYNC","")			
		End If
		
		'Capturing execution end time	
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Check Out",sAction,"","")		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to perform basic check out of object
	Case "CheckOutForm"
		'Setting Change ID
		If sChangeID<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_Common_ObjectCheckOut", "Set",  objCheckOut, "jedt_ChangeID", sChangeID )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to enter value in [ Change ID ] edit box","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Setting Comments
		If sComment<>"" Then
			If Fn_UI_JavaEdit_Operations("RAC_Common_ObjectCheckOut", "Set",  objCheckOut, "jedt_Comments", sComment )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to enter value in Comments edit box","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		'Setting Export Option
		If sExport<>"" Then
			If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectCheckOut", "Set", objCheckOut, "jckb_ExportDtsetOnCheckOut", sExport)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to select checkbox for Export Dataset","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		'Setting Overwrite option
		If sOverwrite<>"" Then
			If Fn_UI_JavaCheckBox_Operations("RAC_Common_ObjectCheckOut", "Set", objCheckOut, "jckb_OverwriteExistingFiles", sOverwrite)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to select checkbox for Overwrite existing files","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If

		'Click on Yes button to Check out Object
		If bFlag= True Then
			If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_OK") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to click on OK button","","","","","")
				Call Fn_ExitTest()
			End If
			'Checking existance of [ Check Out ] warning dialog
			If  Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist", objCheckOut.JavaTree("jtree_ErrorTree"),"","","")=True Then
'				Call Fn_UI_JavaTree_Operations("RAC_Common_ObjectCheckOut","Select",objCheckOut,"jtree_ErrorTree","#0","","")
				objCheckOut.JavaTree("jtree_ErrorTree").DblClick 2,2,"LEFT"
'				JavaWindow("jwnd_DefaultWindow").JavaWindow("jwnd_CheckingOut").JavaTree("jtree_ErrorTree").Select "#0"
				If  Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist", objCheckOut.JavaWindow("jwnd_Warning"),"","","")=True Then
					If Instr(1,Lcase(objCheckOut.JavaWindow("jwnd_Warning").JavaEdit("jedt_ErrorDetails").GetROProperty("value")),Lcase("The instance cannot be saved because it contains at least one attribute that violates a unique attribute rule")) Then
						If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut.JavaWindow("jwnd_Warning"),"jbtn_OK") = False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to click on [ OK ] button od Warning dialog","","","","","")
							Call Fn_ExitTest()
						End If
						If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_OK") = False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to click on OK button","","","","","")
							Call Fn_ExitTest()
						End If
					Else
						If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_OK") = False Then
							Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as unwated error appears while checking out the object","","","","","")
							Call Fn_ExitTest()
						End If
					End If
				End If
			End If
		Else
			If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_Yes") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to check out selected object as fail to click on Yes button","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		If Err.Number < 0 Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Check Out selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		End If				

		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		'Checking existance of [ Check Out ] dialog
		If Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist",objCheckOut,GBL_MICRO_TIMEOUT,"","")=True Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to Check Out selected object as error apeared after check out operation","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully Check Out selected object","","","","DONOTSYNC","")					
		'Capturing execution end time	
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Check Out",sAction,"","")	

	Case "CancelCheckOut"
	
		Set objCheckOut = Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jwnd_CancelCheckOut","")
		
		'Checking existance of [ Check Out ] dialog
		If Fn_UI_Object_Operations("RAC_Common_ObjectCheckOut", "Exist", objCheckOut,GBL_MIN_MICRO_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ " & Cstr(sAction) & " ] operation as [ Cancel Check Out ] dialog does not exist","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click on OK buitton on cancel checkout dialog
		If Fn_UI_JavaButton_Operations("Fn_ObjectCheckOut", "Click", objCheckOut,"jbtn_OK") = False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to [OK] button on Cancel Check Out window","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully closed Cancel Check Out window","","","","","")
		End If
		
End Select

'Releasing Object of [ Check Out ] dialog
Set objCheckOut= Nothing
Set objCheckingOut= Nothing

Function Fn_ExitTest()
	'Releasing Object of [ Check Out ] dialog
	Set objCheckOut= Nothing
	Set objCheckingOut= Nothing
	ExitTest
End Function

