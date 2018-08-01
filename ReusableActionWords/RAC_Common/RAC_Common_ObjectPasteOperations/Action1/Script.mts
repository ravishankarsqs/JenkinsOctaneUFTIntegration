'! @Name RAC_Common_ObjectPasteOperations
'! @Details This actionword is used to perform menu operations in Teamcenter application.
'! @InputParam1. sAction = Action to be performed
'! @InputParam2. sMenuLabel = Menu label tag
'! @Author Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer Kundan Kudale kundan.kudale@sqs.com
'! @Date 25 Mar 2016
'! @Version 1.0
'! @Example  LoadAndRunAction "RAC_Common\RAC_Common_ObjectPasteOperations","RAC_Common_ObjectPasteOperations",OneIteration,"Select","FileNew","RAC_Common_Menu"

Option Explicit

Dim sAction, sInvokeOption, sCopyNodePath, sPasteNodePath

Select Case sAction

	Case "CopyRawMaterialAndPasteInEnggPartRevision"
	
End Select
