'! @Name 			RAC_Common_EditPasteSpecial
'! @Details 		Action word to perform copy paste operations
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sInvokeOption		: Paste Special dialog invoke option
'! @InputParam3 	sPasteAsRelation	: Paste as relation
'! @InputParam4 	sButton				: Button Name
'! @Author 			Shrikant Narkhede Shrikant.Narkhede@sqs.com
'! @Reviewer 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Date 			26 July 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_EditPasteSpecial","RAC_Common_EditPasteSpecial", oneIteration, "PasteSpecial","Relation_HasMirroredHanded",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction, sInvokeOption, sPasteAsRelation, sTempRelation, sButton
Dim objPasteSpecialDialog

GBL_CURRENT_EXECUTABLE_APP="RAC"

'Set parameter values in local variables
sAction= Parameter("sAction")
sInvokeOption = Parameter("sInvokeOption")
sPasteAsRelation = Parameter("sPasteAsRelation")
sButton = Parameter("sButton")

'creating object of Paste Special Dialog
Set objPasteSpecialDialog=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","jdlg_PasteSpecial","")

If sPasteAsRelation <> "" Then
	sTempRelation = sPasteAsRelation
	sPasteAsRelation = Fn_FSOUtil_XMLFileOperations("getvalue","RAC_PasteSpecialRelations_APL",sPasteAsRelation,"")
	IF sPasteAsRelation = False Then
		sPasteAsRelation = sTempRelation
	End IF 
End IF

'Invoking [ Paste Special ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","EditPasteSpecial"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_EditPasteSpecial"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Select Case sAction
	Case "PasteSpecial"
		'Verify existence of list of values dropdown
		If Fn_UI_Object_Operations("RAC_Common_EditPasteSpecial", "Exist", objPasteSpecialDialog.JavaList("jlst_AddAs"),"","","") Then			
			'Verify existence of value
			If Fn_UI_JavaList_Operations("RAC_Common_EditPasteSpecial","Exist",objPasteSpecialDialog,"jlst_AddAs",sPasteAsRelation, "", "")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify relation [ " & Cstr(sPasteAsRelation) & " ] as it does not exist in listbox","","","","","")
				Call Fn_ExitTest()
			End If
			
			'Select the value
			If Fn_UI_JavaList_Operations("RAC_Common_EditPasteSpecial","select",objPasteSpecialDialog,"jlst_AddAs",sPasteAsRelation, "", "") = False Then					
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select relation [ " & Cstr(sPasteAsRelation) & " ] from listbox","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully selected relation [ " & Cstr(sPasteAsRelation) & " ]","","","","","")
			End If
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to verify existence of Paste Special dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		'Click On OK Button
		IF Fn_UI_JavaButton_Operations("RAC_Common_EditPasteSpecial", "Click", objPasteSpecialDialog,"jbtn_OK")=False then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform [ Paste Spacial ] operation as fail to click on [ OK ] button of Paste Special dialog","","","","","")
			Call Fn_ExitTest()
		End IF
		Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End Select

'Releasing object of New BusinessObject dialog
Set objPasteSpecialDialog =Nothing

Function Fn_ExitTest()
	Set objPasteSpecialDialog =Nothing
	ExitTest
End Function
