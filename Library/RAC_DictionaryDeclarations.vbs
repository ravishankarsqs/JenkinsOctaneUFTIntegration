Option Explicit

'Declaring variables
Dim dictWorkflowProcessInfo
Dim dictProperties                                    
Dim dictEditProperties
Dim dictItemInfo
Dim dictPartInfo
Dim dictDesignInfo
Dim dictErrorMessageInfo
Dim dictAssignParticipantInfo
Dim dictDatasetInfo
Dim dictSummaryTabInfo
Dim dictReviseInfo
Dim dictViewerTabInfo
Dim dictUserSessionInfo
Dim dictChangeInfo
Dim dictImpactAnalysisTabInfo
Dim dictFormInfo
Dim dictBusinessObjectInfo
Dim dictBaselineInfo
Dim dictSaveAsInfo
Dim dictMassUpdateInfo
Dim dictGenerateReportInformation
Dim dictDeriveChangeInfo
Dim dictRemoteExportInfo
Dim dictAssignResponsiblePartyInfo
Dim dictDispatcher
Dim dictProgramProjectInfo
Dim dictJTPreviewTabInfo

'to store workflow process infomation
Set dictWorkflowProcessInfo=CreateObject("Scripting.Dictionary")
'to store object properties infomation
Set dictProperties=CreateObject("Scripting.Dictionary")
'to store object properties infomation
Set dictEditProperties=CreateObject("Scripting.Dictionary")
'to store item creation additional infomation
Set dictItemInfo = CreateObject("Scripting.Dictionary")
'to store part creation additional infomation
Set dictPartInfo = CreateObject("Scripting.Dictionary")
'to store Design creation additional infomation
Set dictDesignInfo = CreateObject("Scripting.Dictionary")
'to store error message additional infomation
Set dictErrorMessageInfo = CreateObject("Scripting.Dictionary")
'to store assing participants infomation
Set dictAssignParticipantInfo=CreateObject("Scripting.Dictionary")
'to store dataset infomation
Set dictDatasetInfo=CreateObject("Scripting.Dictionary")
'to store summary tab infomation
Set dictSummaryTabInfo=CreateObject("Scripting.Dictionary")
'to store revise
Set dictReviseInfo=CreateObject("Scripting.Dictionary")
'to store viewer tab information
Set dictViewerTabInfo=CreateObject("Scripting.Dictionary")
'to store User Session information
Set dictUserSessionInfo=CreateObject("Scripting.Dictionary")
'to store change object information
Set dictChangeInfo=CreateObject("Scripting.Dictionary")
'to store information while performing operation on impact analysis tab
Set dictImpactAnalysisTabInfo=CreateObject("Scripting.Dictionary")
'to store form creation additional infomation
Set dictFormInfo = CreateObject("Scripting.Dictionary")
'to store Business Object creation additional infomation
Set dictBusinessObjectInfo = CreateObject("Scripting.Dictionary")
'to store Baseline related information
Set dictBaselineInfo = CreateObject("Scripting.Dictionary")
'to store Save As object information
Set dictSaveAsInfo=CreateObject("Scripting.Dictionary")
'to store mass update information
Set dictMassUpdateInfo=CreateObject("Scripting.Dictionary")
'to store report generation information
Set dictGenerateReportInformation=CreateObject("Scripting.Dictionary")
'to store Derive change object information
Set dictDeriveChangeInfo=CreateObject("Scripting.Dictionary")
'to store Remote Export information
Set dictRemoteExportInfo=CreateObject("Scripting.Dictionary")
'to store Assign Responsible Party information
Set dictAssignResponsiblePartyInfo=CreateObject("Scripting.Dictionary")
'to store dispatcher related information
Set dictDispatcher=CreateObject("Scripting.Dictionary")
'to store Program Wrapper information
Set dictProgramProjectInfo=CreateObject("Scripting.Dictionary")
Set dictJTPreviewTabInfo=CreateObject("Scripting.Dictionary")