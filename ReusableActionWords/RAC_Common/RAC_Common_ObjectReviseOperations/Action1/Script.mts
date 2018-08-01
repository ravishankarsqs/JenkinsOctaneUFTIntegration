'! @Name 			RAC_Common_ObjectReviseOperations
'! @Details 		Action word use to perform operation on Objects Revise dialog
'! @InputParam1 	sAction 		: Action to be performed e.g. autobasicrevise
'! @InputParam2 	sInvokeOption 	: Method to invoke revise dialog e.g. menu
'! @InputParam4 	sButton		 	: Button Name
'! @InputParam5 	dictReviseInfo 	: External parameter to pass revise object additional information
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			13 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_ObjectReviseOperations","RAC_Common_ObjectReviseOperations",OneIteration,"autobasicrevise","Menu",""

Option Explicit
Err.Clear

'Declaring varaibles	
Dim sAction,sInvokeOption,sButton
Dim sID,sName,sRevision,sPerspective,sBasedOn
Dim iReviseObjectCount
Dim objRevise
Dim aRevision
Dim iCounter
Dim objDescription,objChildObjects

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction=Parameter("sAction")
sInvokeOption=Parameter("sInvokeOption")
sButton = Parameter("sButton")

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Get active perspective name
sPerspective=Fn_RAC_GetActivePerspectiveName("")

'Creating object of Revise dialog
Select Case LCase(sPerspective)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "myteamcenter","","structuremanager"
		Set objRevise=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_Revise","")
End Select

'inoke Revise dialog
Select Case LCase(sInvokeOption)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileRevise"
		'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'Use this invoke option when user wants to invoke Revise dialog from outside function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_ObjectReviseOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of revise dialog
If Fn_UI_Object_Operations("RAC_Common_ObjectReviseOperations", "Exist", objRevise, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ Revise ] dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting Revise object count
If Lcase(sAction)= Lcase("autobasicrevise") OR Lcase(sAction)= Lcase("autobasicrevisewithoutclose") OR Lcase(sAction)= Lcase("basicrevise") OR Lcase(sAction)= Lcase("autobasicrevisewithautopopulatedfields")  OR Lcase(sAction)= Lcase("basicrevisewithoutclose") Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectCount","","")
	iReviseObjectCount=Fn_CommonUtil_DataTableOperations("GetValue","RACReviseObjectCount","","")
	If iReviseObjectCount="" Then
		iReviseObjectCount=1
	Else
		iReviseObjectCount=iReviseObjectCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACReviseObjectCount",iReviseObjectCount,"")
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","Object Revise Operations",sAction,"","")

Select Case Lcase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to revise obejct
	Case "autobasicrevise","autobasicrevisewithoutclose","basicrevise","basicrevisewithoutclose"
		'Setting revision		
		If Lcase(sAction)="basicrevise" or Lcase(sAction)="basicrevisewithoutclose" Then
			
			sRevision=dictReviseInfo("Revision")
			dictReviseInfo("Revision")=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ReviseObjectValues_APL",dictReviseInfo("Revision"),""))
			If cStr(dictReviseInfo("Revision"))="False" Then
				dictReviseInfo("Revision")=sRevision
			End If
			sRevision=""
			
			If Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations","Set",objRevise,"jedt_Revision",dictReviseInfo("Revision"))=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object as fail to set revision id value","","","","","")	
				Call Fn_ExitTest()
			End If
			'Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		Else
			IF Fn_UI_Object_Operations("RAC_Common_ObjectReviseOperations","GetROProperty",objRevise.JavaButton("jbtn_Assign"),"","enabled","")=1 Then
				If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_Assign")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object as fail to click on [ Assign ] button to assign new revision","","","","","")	
					Call Fn_ExitTest()
				End If							
			End IF
		End If		
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Getting New Revision
		If objRevise.JavaEdit("jedt_Revision").Exist(1) Then
			sRevision=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_Revision", "")
		ElseIf objRevise.JavaEdit("jedt_Revision2").Exist(0) Then
			sRevision=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_Revision2", "")		
		End If		
		sID=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",objRevise,"jedt_ID", "")
		sName=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_Name", "")
		sBasedOn=objRevise.JavaStaticText("jstx_BasedOnInformation").GetROProperty("label")
		
		'Clicking on Finish Button
		objRevise.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 60000
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations","click",objRevise,"jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object as fail to click on [ Finish ] button","","","","","")	
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Clicking on Close Button
		If sAction<>"autobasicrevisewithoutclose" and Lcase(sAction)<>"basicrevisewithoutclose" and Lcase(sAction)<>"basicrevisecontroldocumentwithoutclose" Then
			If Fn_UI_Object_Operations("RAC_Common_ObjectReviseOperations", "Exist", objRevise.JavaButton("jbtn_Close"),GBL_MICRO_TIMEOUT,"","") Then
				If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_Close")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object as fail to click on [ Close ] button","","","","","")	
					Call Fn_ExitTest()
				End If
				Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
			End If
		End If
						
		sBasedOn=Mid(sBasedOn,1,Instr(1,sBasedOn,"(Type:")-1)
		sBasedOn=Split(Trim(sBasedOn),sName)
		sName=sName & sBasedOn(1)
		
		DataTable.SetCurrentRow iReviseObjectCount
		'Setting Revision ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectRevisionID","","")					
		DataTable.Value("RACReviseObjectRevisionID","Global")= sRevision
		
		'Setting ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectID","","")					
		DataTable.Value("RACReviseObjectID","Global")= sID
		
		'Setting Revision Name in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectRevisionName","","")					
		DataTable.Value("RACReviseObjectRevisionName","Global")= sName	
		
		'Setting ReviseObject Revision Node value in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectRevisionNode","","")
		If Lcase(sAction)="autobasicrevisecontroldocument" Or Lcase(sAction)="autobasicreviseprogram" Or Lcase(sAction)="basicrevisecontroldocument"Then
			DataTable.Value("RACReviseObjectRevisionNode","Global")=DataTable.Value("RACReviseObjectID","Global") & "/" & DataTable.Value("RACReviseObjectRevisionID","Global") & ";1-" & DataTable.Value("RACReviseObjectRevisionName","Global")
		Else
			DataTable.Value("RACReviseObjectRevisionNode","Global")=DataTable.Value("RACReviseObjectID","Global") & "/" & DataTable.Value("RACReviseObjectRevisionID","Global") & "-" & DataTable.Value("RACReviseObjectRevisionName","Global")
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Revise Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			If sAction="autobasicrevisewithoutclose" or Lcase(sAction)="basicrevisewithoutclose" or Lcase(sAction)="basicrevisecontroldocumentwithoutclose" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully set Revision Id [ " & Cstr(Datatable.Value("RACReviseObjectRevisionID", "Global")) & " ] on revise dialog for selected object and click on [ Finish ] button","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully revised selected object to Revision Id [ " & Cstr(Datatable.Value("RACReviseObjectRevisionID", "Global")) & " ]","","","","DONOTSYNC","")
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to click on specific button of revise dialog
	Case "clickbutton"
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of Revise dialog","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Revise Operations",sAction,"Button Name",sButton)
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully click on [ " & Cstr(sButton) & " ] button of Revise dialog","","","","","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to revise obejct
	Case "autobasicrevisewithautopopulatedfields"
		'Getting New Revision
		sRevision=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_Revision", "")
		sID=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",objRevise,"jedt_ID", "")
		sName=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_Name", "")
		
		'Clicking on Finish Button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations","click",objRevise,"jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object as fail to click on [ Finish ] button","","","","","")	
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		
		'Clicking on Close Button
		If Fn_UI_Object_Operations("RAC_Common_ObjectReviseOperations", "Exist", objRevise.JavaButton("jbtn_Close"),"","","") Then
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_Close")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object as fail to click on [ Close ] button","","","","","")	
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)	
		End If
		
		DataTable.SetCurrentRow iReviseObjectCount
		'Setting Revision ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectRevisionID","","")					
		DataTable.Value("RACReviseObjectRevisionID","Global")= sRevision
		
		'Setting ID in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectID","","")					
		DataTable.Value("RACReviseObjectID","Global")= sID
		
		'Setting Revision Name in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectRevisionName","","")					
		DataTable.Value("RACReviseObjectRevisionName","Global")= sName	
		
		'Setting ReviseObject Revision Node value in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACReviseObjectRevisionNode","","")
		If Lcase(sAction)="autobasicrevisecontroldocumentwithautopopulatedfields" Then
			DataTable.Value("RACReviseObjectRevisionNode","Global")=DataTable.Value("RACReviseObjectID","Global") & "/" & DataTable.Value("RACReviseObjectRevisionID","Global") & ";1-" & DataTable.Value("RACReviseObjectRevisionName","Global")
		Else
			DataTable.Value("RACReviseObjectRevisionNode","Global")=DataTable.Value("RACReviseObjectID","Global") & "/" & DataTable.Value("RACReviseObjectRevisionID","Global") & "-" & DataTable.Value("RACReviseObjectRevisionName","Global")
		End If
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Revise Operations",sAction,"","")
		If Err.Number<>0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to revise selected object due to error [ " & Cstr(Err.Description) & " ]","","","","DONOTSYNC","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully revised selected object to Revision Id [ " & Cstr(Datatable.Value("RACReviseObjectRevisionID", "Global")) & " ]","","","","DONOTSYNC","")
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify revision field is auto populate on revise dialog
	Case "verifyrevisionisautopopulate"
	    sRevision=Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_Revision", "")
		If sRevision = "" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as revision field is not auto populate on [ Revise ] dialog","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Revise Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified revision field is auto populate shows Revision [ " & Cstr(sRevision) & "] on [ Revise ] dialog","","","","DONOTSYNC","")
		End If
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr("jbtn_" & sButton) & " ] button of Revise dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify id field is auto populate on revise dialog
	Case "verifyidisautopopulate"
		If Fn_UI_JavaEdit_Operations("RAC_Common_ObjectReviseOperations", "GetText",  objRevise,"jedt_ID", "") = "" Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Verification fail as ID field is not auto populate on [ Revise ] dialog","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","Object Revise Operations",sAction,"","")
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified ID field is auto populate on [ Revise ] dialog","","","","DONOTSYNC","")
		End If
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr("jbtn_" & sButton) & " ] button of Revise dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	'Case to verify revision id does not exist
	Case "verifyrevisionidnotexist"
		sRevision=dictReviseInfo("Revision")
		dictReviseInfo("Revision")=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ReviseObjectValues_APL",dictReviseInfo("Revision"),""))
		If cStr(dictReviseInfo("Revision"))="False" Then
			dictReviseInfo("Revision")=sRevision
		End If
		sRevision=""
		
		'Clicking on Revision Id Dropdown Button
		If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations","click",objRevise,"jbtn_RevisionIdDropdown")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Revision Id Dropdown ] button of Revise dialog","","","","","")	
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
		aRevision=Split(dictReviseInfo("Revision"),"~")
		For iCounter=0 to Ubound(aRevision)
			Set objDescription = Description.Create()
			objDescription("Class Name").value = "JavaStaticText"
			objDescription("label").value = aRevision(iCounter) & ", "
			Set objChildObjects = objNewDataset.ChildObjects(objDescription)
			If objChildObjects.count <> 0 Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : Successfully verified revision id [ " & Cstr(aRevision(iCounter)) & " ] does not available on revise id list of revise dialog","","","","DONOTSYNC","")
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as revision id [ " & Cstr(aRevision(iCounter)) & " ] available on revise id list of revise dialog","","","","","")
				Call Fn_ExitTest()
			End If		
			Set objChildObjects = Nothing
			Set objDescription =Nothing
		Next
		
		objRevise.JavaEdit("jedt_Revision").Click 1,1
		wait 1
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_ObjectReviseOperations", "Click", objRevise,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr("jbtn_" & sButton) & " ] button of Revise dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_DEFAULT_SYNC_ITERATIONS)
		End If
End Select

DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

'Releasing object
Set objRevise=Nothing

Function Fn_ExitTest()
	'Releasing object
	Set objRevise=Nothing
	ExitTest
End Function

