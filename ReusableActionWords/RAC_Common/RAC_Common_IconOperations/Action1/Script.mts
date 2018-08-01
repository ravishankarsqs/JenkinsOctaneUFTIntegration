'! @Name 			RAC_Common_IconOperations
'! @Details 		Actionword To perform operations on Icons\Images in teamcenter
'! @InputParam1 	sAction 		: String to indicate what action is to be performed on Icons\Images in teamcenter
'! @InputParam2 	sIconName 		: Icon name
'! @InputParam3 	sIconInstance	: Icon Instance number
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			09 May 2017
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_IconOperations","RAC_Common_IconOperations",oneIteration,"VerifyExist","Flag",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sIconName,sIconInstance
Dim bFlag
Dim objInsightObject,sLogStatement

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get parameter values in local variables
sAction = Parameter("sAction")
sIconName = Parameter("sIconName")
sIconInstance = Parameter("sIconInstance")

bFlag = False

'creating object of Icon
Select Case sIconName
	Case "PartRevisionOrangeSpot"
		sLogStatement="Orange Spot"
	Case "CATPartDatasetIcon"
		sLogStatement="CAT Part Dataset Gear"
	Case "CATDrawingDatasetIcon"
		sLogStatement="CAT Drawing Dataset Sheet"
	Case "JTDatasetIcon"
		sLogStatement="Direct Model Dataset"	
	Case "EngineeredDrawingIcon"
		sLogStatement="CATIA File Logo"	
	Case "Flag"
		sLogStatement="Flag"
	Case "EngineeredPartRevisionIcon"
		sLogStatement="Engineered Part Revision WIP Icon"	
	Case "PartRevisionGreenSpot"
		sLogStatement="Green Spot"
	Case "BaselineFlagIcon"
		sLogStatement="Baseline Flag Icon"
	Case "SupportDesignRevisionIcon"
		sLogStatement="Support Design Revision Icon"
	Case "RawMaterialRevisionGreenSpot"
		sLogStatement="Raw Material Revision Green Spot"
	Case "RawMaterialRevisionRedSpot"
		sLogStatement="Raw Material Revision Red Spot"
	Case "DFFrozenGreenTick"
		sLogStatement="Design Freeze Green tick frozen status"
	Case "CheckOutFlag"
		sLogStatement="Check Out Flag"
	Case "CNFrozenGreenTick"
		sLogStatement="Change Notice Green tick frozen status"
	Case "DesignFreezeErrorCross"
		sLogStatement="Design Freeze Error red cross"
	Case "ChangeNoticeErrorCross"
		sLogStatement="Change Notice Error red cross"
	Case Else
		sLogStatement=sIconName
End Select

Set objInsightObject=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_OU_OR","iobj_" & sIconName,"")

If sIconInstance<>"" Then
	sIconInstance=Cint(sIconInstance)-1
	objInsightObject.SetTOProperty "index",Cint(sIconInstance)
End If

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_IconOperations"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

Select Case sAction
	Case "VerifyExist","VerifyNonExist"
		bFlag = True
		If Fn_UI_Object_Operations("RAC_Common_IconOperations","Exist", objInsightObject,"","","") = False Then
			bFlag = False
		End If

		If bFlag = False Then
			If sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sLogStatement) & " ] icon does not exist","","","","","")
				Call Fn_ExitTest()
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sLogStatement) & " ] icon does not exist","","","","DONOTSYNC","")
			End If
		Else
			If sAction="VerifyExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(sLogStatement) & " ] icon exist","","","","DONOTSYNC","")
			ElseIf sAction="VerifyNonExist" Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(sLogStatement) & " ] icon exist","","","","","")
				Call Fn_ExitTest()
			End If
		End If
End Select

Set objInsightObject=Nothing

Function Fn_ExitTest()	
	'Releasing all objects
	Set objInsightObject=Nothing
	ExitTest
End Function
