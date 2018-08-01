'! @Name 			RAC_Common_SetResetPerspective
'! @Details 		To set and reset a perspective in Teamcenter
'! @InputParam1 	sPerspective	: perspective name
'! @InputParam2		bSetFlag		: Perspective set flag [ Option ]
'! @InputParam3		bResetFlag		: Perspective reset flag [ Option ]
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			26 Mar 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_SetResetPerspective","RAC_Common_SetResetPerspective",OneIteration,"My Teamcenter",True,True

Option Explicit
Err.Clear

'Declaring varaibles
Dim sPerspective
Dim bSetFlag,bResetFlag

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Get action parameters in local variables
sPerspective = Parameter("sPerspective")
bSetFlag = Parameter("bSetFlag")
bResetFlag = Parameter("bResetFlag")

'If Fn_RAC_GetActivePerspectiveName("getname")="Getting Started" Then
'	bResetFlag=False
'End If

'Setting perspective
If sPerspective<>"" Then
	If bSetFlag<>False Then
		If sPerspective<>Fn_RAC_GetActivePerspectiveName("getname") Then
			LoadAndRunAction "RAC_Common\RAC_Common_SetPerspective","RAC_Common_SetPerspective",OneIteration,sPerspective
			'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If		
	End If	
End If
'Resetting perspective
If bResetFlag<>False Then
	LoadAndRunAction "RAC_Common\RAC_Common_ResetPerspective","RAC_Common_ResetPerspective",OneIteration
	'Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
End If

