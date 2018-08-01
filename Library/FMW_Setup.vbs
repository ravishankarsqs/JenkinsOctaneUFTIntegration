Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Function List
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'			Function Name								|     Developer					|		Date	|	Comment
'- - - - - - - - - - - - - - - - - - - - - - - - - - - -| - - - - - - - - - - - - - - - | - - - - - - - | - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'001. Fn_Setup_GetAutomationXMLPath						|	sandeep.navghane@sqs.com	|	16-Jan-2015	|	Function used to get automation XML file path 
'002. Fn_Setup_GetAutomationFolderPath					|	sandeep.navghane@sqs.com	|	16-Jan-2015	|	Function used to get automation folder path 
'003. Fn_Setup_ReporterFilter							|	vrushali.sahare@sqs.com		|	24-Feb-2016	|	Function Used to set Reporter filter
'004. Fn_Setup_ClearRACCache							|	vrushali.sahare@sqs.com		|	24-Feb-2016	|	Function Used to clear cache
'005. Fn_Setup_GetTestUserDetailsFromExcelOperations	|	vrushali.sahare@sqs.com  	|	24-Feb-2016	|	Function used to get user details from specified excel file
'006. Fn_Setup_SetActionIterationMode					|	sandeep.navghane@sqs.com	|	10-Mar-2016	|	Function used to set test case run iteration mode
'007. Fn_Setup_GenerateObjectInformation				|	sandeep.navghane@sqs.com	|	08-Jul-2016	|	Function is used to generate information of application objects
'008. Fn_Setup_DeleteCacheFiles							|	sandeep.navghane@sqs.com	|	21-Sep-2017	|	Function used to delete cache files
'009. Fn_Setup_GetNXTestData							|	sandeep.navghane@sqs.com	|	30-Oct-2017	|	Function is used to get information of NX test data Details
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Setup_GetAutomationXMLPath
'
'Function Description	 :	Function used to get automation XML file path 
'
'Function Parameters	 :  1.sAutomationXMLName : Automation XML Name
'
'Function Return Value	 : 	Automation XML file path \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Automation XML file should exist
'
'Function Usage		     :  bReturn = Fn_Setup_GetAutomationXMLPath("EnvironmentVariables")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
'Replacement for function : Fn_LogUtil_GetXMLPath,Fn_SetEnvValue ' Delete this comment once implementation is completed
Public Function Fn_Setup_GetAutomationXMLPath(sAutomationXMLName)
	'Initially set function return value as False
	Fn_Setup_GetAutomationXMLPath=False
	Environment.Value("AutomationDirPath") = Fn_CommonUtil_EnvironmentVariablesOperations("Get","User","AutomationDir","")
	
	Select Case sAutomationXMLName
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get EnvironmentVariables.xml file path
		Case "EnvironmentVariables"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\SetupXML\EnvironmentVariables.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Common_Menu.xml file path
		Case "RAC_Common_Menu"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\MenuXML\RAC_Common_Menu.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX_Common_Menu.xml file path
		Case "NX_Common_Menu"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\MenuXML\NX_Common_Menu.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_StructureManager_Menu.xml file path
		Case "RAC_StructureManager_Menu"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\MenuXML\RAC_StructureManager_Menu.xml"		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Common_PopupMenu.xml file path
		Case "RAC_Common_PPM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\PopupMenuXML\RAC_Common_PPM.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Common_Toolbar.xml file path
		Case "RAC_Common_TLB"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ToolBarXML\RAC_Common_TLB.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Common_FU_OR.xml file path
		Case "RAC_Common_FU_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_Common_FU_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Common_OU_OR.xml file path
		Case "RAC_Common_OU_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_Common_OU_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_LoginUtil_OR.xml file path
		Case "RAC_LoginUtil_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_LoginUtil_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_MyTeamcenter_OR.xml file path
		Case "RAC_MyTeamcenter_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_MyTeamcenter_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX_LoginUtil_OR.xml file path
		Case "NX_LoginUtil_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\NX_LoginUtil_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX_Common_OR.xml file path
		Case "NX_Common_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\NX_Common_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX_ErrorMessage_OR.xml file path
		Case "NX_ErrorMessage_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\NX_ErrorMessage_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Search_OR.xml file path
		Case "RAC_Search_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_Search_OR.xml"			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_StructureManager_OR.xml file path
		Case "RAC_StructureManager_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_StructureManager_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ChangeManager_OR.xml file path
		Case "RAC_ChangeManager_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_ChangeManager_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ErrorMessage_OR.xml file path
		Case "RAC_ErrorMessage_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_ErrorMessage_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_MyWorklist_OR.xml file path
		Case "RAC_MyWorklist_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_MyWorklist_OR.xml"		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Project_OR.xml file path
		Case "RAC_Project_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_Project_OR.xml"		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get AWC_Common_OR.xml file path
		Case "AWC_Common_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\AWC_Common_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get AWC_LoginUtil_OR.xml file path
		Case "AWC_LoginUtil_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\AWC_LoginUtil_OR.xml"		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX_ErrorMessage_ERM.xml file path
		Case "NX_ErrorMessage_ERM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ErrorMessageXML\NX_ErrorMessage_ERM.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ErrorMessage_ERM.xml file path
		Case "RAC_ErrorMessage_ERM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ErrorMessageXML\RAC_ErrorMessage_ERM.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NavigationTreeNodeTypeInformation_APL.xml file path
		Case "RAC_NavigationTreeNodeTypeInformation_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NavigationTreeNodeTypeInformation_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_BaselineValues_APL.xml file path
		Case "RAC_BaselineValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_BaselineValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewChangeValues_APL.xml file path
		Case "RAC_NewChangeValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewChangeValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get AWC_NewEventValues_APL.xml file path
		Case "AWC_NewEventValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\AWC_NewEventValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewDesignValues_APL.xml file path
		Case "RAC_NewDesignValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewDesignValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewItemValues_APL.xml file path
		Case "RAC_NewItemValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewItemValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewPartValues_APL.xml file path
		Case "RAC_NewPartValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewPartValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get AWC_NewProgramValues_APL.xml file path
		Case "AWC_NewProgramValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\AWC_NewProgramValues_APL.xml"				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ObjectPropertiesValues_APL.xml file path
		Case "RAC_ObjectPropertiesValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_ObjectPropertiesValues_APL.xml"								
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_RevisionRuleValues_APL.xml file path
		Case "RAC_RevisionRuleValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_RevisionRuleValues_APL.xml"											
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ProgramValues_PRE.xml file path
		Case "RAC_ProgramValues_PRE"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\PrerequisiteInformationXML\RAC_ProgramValues_PRE.xml"								
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get BasicModel_InputInformation_TD.xml file path
		Case "BasicModel_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\BasicModel_InputInformation_TD.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get BlockCreate_InputInformation_TD.xml file path
		Case "BlockCreate_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\BlockCreate_InputInformation_TD.xml"			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Model_InputInformation_TD.xml file path
		Case "Model_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Model_InputInformation_TD.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Model_OutputInformation_TD.xml file path
		Case "Model_OutputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Model_OutputInformation_TD.xml"						
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ProcessTemplateName_WF.xml file path
		Case "RAC_ProcessTemplateName_WF"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\WorkflowXML\RAC_ProcessTemplateName_WF.xml"										
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_Common_OR.xml file path
		Case "CATIA_Common_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\CATIA_Common_OR.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_Common_OR.xml file path
		Case "CATIA_Command"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\CommandXML\CATIA_Command.xml"	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_Common_OR.xml file path
		Case "CATIA_CommandInformation"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\Utilities\FunctionUtilityVBS\CATIA_CommandInformation.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_Common_OR.xml file path
		Case "CATIA_Common_TLB"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ToolBarXML\CATIA_Common_TLB.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_Common_OR.xml file path
		Case "CATIA_Common_Menu"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\MenuXML\CATIA_Common_Menu.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_TeamcenterSaveManagerValues_APL.xml file path
		Case "CATIA_TeamcenterSaveManagerValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\CATIA_TeamcenterSaveManagerValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewDatasetValues_APL.xml file path
		Case "RAC_NewDatasetValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewDatasetValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewFormValues_APL.xml file path
		Case "RAC_NewFormValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewFormValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_ReviseObjectValues_APL.xml file path
		Case "RAC_ReviseObjectValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_ReviseObjectValues_APL.xml"				
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_Common_PPM.xml file path
		Case "RAC_Common_PPM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\PopUpMenuXML\RAC_Common_PPM.xml"	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_Common_PPM.xml file path
		Case "CATIA_Common_PPM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\PopUpMenuXML\CATIA_Common_PPM.xml"		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_StructureManager_TLB.xml file path
		Case "RAC_StructureManager_TLB"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ToolBarXML\RAC_StructureManager_TLB.xml"	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA environment values
		Case "CATIA_EnvironmentValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\CATIA_EnvironmentValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA drawing template values
		Case "CATIA_DrawingTemplate_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\CATIA_DrawingTemplate_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get CATIA_ErrorMessage_ERM.xml file path
		Case "CATIA_ErrorMessage_ERM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ErrorMessageXML\CATIA_ErrorMessage_ERM.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_NewProgramValues_APL.xml file path
		Case "RAC_NewProgramValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_NewProgramValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_SearchCriteriaValues_APL.xml file path
		Case "RAC_SearchCriteriaValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_SearchCriteriaValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get RAC_PasteSpecialRelations_APL.xml file path
		Case "RAC_PasteSpecialRelations_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\RAC_PasteSpecialRelations_APL.xml"			
		'Case to get ApplyMaterial_InputInformation_TD.xml file path
		Case "ApplyMaterial_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\ApplyMaterial_InputInformation_TD.xml"
		'Case to get Drawing_InputInformation_TD.xml file path
		Case "Drawing_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Drawing_InputInformation_TD.xml"
		'Case to get Drawing_OutputInformation_TD.xml file path
		Case "Drawing_OutputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Drawing_OutputInformation_TD.xml"
		'Case to get DrawingView_InputInformation_TD.xml file path
		Case "DrawingView_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\DrawingView_InputInformation_TD.xml"
		'Case to get SaveAs_InputInformation_TD.xml file path
		Case "SaveAs_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\SaveAs_InputInformation_TD.xml"
		'Case to get SaveAs_OutputInformation_TD.xml file path
		Case "SaveAs_OutputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\SaveAs_OutputInformation_TD.xml"
		'Case to get NX_ErrorMessage_ERM.xml file path
		Case "NX_ErrorMessage_ERM"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\NX\ErrorMessageXML\NX_ErrorMessage_ERM.xml"
		'Case to get Assembly_InputInformation_TD.xml file path
		Case "Assembly_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Assembly_InputInformation_TD.xml"	
		'Case to get Assembly_OutputInformation_TD.xml file path
		Case "Assembly_OutputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Assembly_OutputInformation_TD.xml"	
		'Case to get NX_ObjectPropertiesValues_APL.xml file path
		Case "NX_ObjectPropertiesValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\NX_ObjectPropertiesValues_APL.xml"				
		'Case to get ImportAssembly_InputInformation_TD.xml file path
		Case "ImportAssembly_InputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\ImportAssembly_InputInformation_TD.xml"	
		'Case to get GetProperties_OutputInformation_TD.xml file path
		Case "GetProperties_OutputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\GetProperties_OutputInformation_TD.xml"	
		'Case to get Model_OutputInformation_TD.xml file path
		Case "Model_OutputInformation_TD"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\TestData\NX\Journal\Model_OutputInformation_TD.xml"	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "AWC_Object_Values"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\AWC_Object_Values.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get TC environment values
		Case "TC_EnvironmentValues_APL"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ApplicationInformationXML\TC_EnvironmentValues_APL.xml"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Organization OR object values
		Case "RAC_Organization_OR"
			Fn_Setup_GetAutomationXMLPath = Environment.Value("AutomationDirPath") & "\AutomationXML\ObjectRepositoryXML\RAC_Organization_OR.xml"
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Setup_GetAutomationFolderPath
'
'Function Description	 :	Function used to get automation folder path 
'
'Function Parameters	 :  1.sFolderName : Folder Name
'
'Function Return Value	 : 	Automation folder path \ False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	Automation folder should exist
'
'Function Usage		     :  bReturn = Fn_Setup_GetAutomationFolderPath("TestData")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Jan-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_GetAutomationFolderPath(sFolderName)
	Dim sComputerUserName,sLastModifiedFolder
	'Initially set function return value as False
	Fn_Setup_GetAutomationFolderPath=False
	Environment.Value("AutomationDirPath") = Fn_CommonUtil_EnvironmentVariablesOperations("Get","User","AutomationDir","")
	
	Select Case sFolderName
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Test data folder path
		Case "TestData"
			Fn_Setup_GetAutomationFolderPath=Environment.Value("AutomationDirPath") & "\TestData"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX macro folder path
		Case "Macro"
			Fn_Setup_GetAutomationFolderPath=Environment.Value("AutomationDirPath") & "\TestData\NX\Macro"	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get NX folder path
		Case "NX"
			Fn_Setup_GetAutomationFolderPath=Environment.Value("AutomationDirPath") & "\TestData\NX"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		'Case to get NX Journal folder path
		Case "NXJournal"
			Fn_Setup_GetAutomationFolderPath=Environment.Value("AutomationDirPath") & "\TestData\NX\Journal"						
				' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		'Case to get tcic_tmp_catuii or  tcic_tmp_Export folder path			
		Case "tcic_tmp_catuii","tcic_tmp_Export","tcic_tmp_Import","tcic_tmp_tmp","tcic_tmp"
			sComputerUserName=Fn_CommonUtil_LocalMachineOperations("getcurrentloginusername","")
			sLastModifiedFolder=Fn_FSOUtil_FolderOperations("getlastmodifiedsubfolder", "C:\tcplm\temp\tcic_tmp\" & sComputerUserName, "", "","")
			
			If sLastModifiedFolder<>False Then
				If sFolderName="tcic_tmp_catuii" Then
					Fn_Setup_GetAutomationFolderPath="C:\tcplm\temp\tcic_tmp\" & sComputerUserName & "\" & sLastModifiedFolder & "\catuii"
				ElseIf sFolderName="tcic_tmp_Export" Then	
					Fn_Setup_GetAutomationFolderPath="C:\tcplm\temp\tcic_tmp\" & sComputerUserName & "\" & sLastModifiedFolder & "\Export"
				ElseIf sFolderName="tcic_tmp_Import" Then	
					Fn_Setup_GetAutomationFolderPath="C:\tcplm\temp\tcic_tmp\" & sComputerUserName & "\" & sLastModifiedFolder & "\Import"
				ElseIf sFolderName="tcic_tmp_tmp" Then	
					Fn_Setup_GetAutomationFolderPath="C:\tcplm\temp\tcic_tmp\" & sComputerUserName & "\" & sLastModifiedFolder & "\tmp"
				ElseIf sFolderName="tcic_tmp" Then	
					Fn_Setup_GetAutomationFolderPath="C:\tcplm\temp\tcic_tmp" 						
				End If
			Else
				Fn_Setup_GetAutomationFolderPath="INVALID FOLDER PATH"
			End If
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_ReporterFilter
'
'Function Description	 :	Function used to set Reporter filter
'
'Function Parameters	 :   1.sAction: Name of the action to be performed
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	Call Fn_Setup_ReporterFilter("DisableAll")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  24-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_ReporterFilter(sAction)
	'Initially set function return value as False
	Fn_Setup_ReporterFilter = False
	
	Select Case lcase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to enable all
		Case "enableall"
			Reporter.Filter = rfEnableAll
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to enable errors and warnings
		Case "enableerrorsandwarnings"
			Reporter.Filter = rfEnableErrorsAndWarnings
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to enable errors only
		Case "enableerrorsonly"
			Reporter.Filter = rfEnableErrorsOnly
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to disable all
		Case "disableall"
			Reporter.Filter = rfDisableAll
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Exit Function
	End Select
	Fn_Setup_ReporterFilter = True
End Function

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_ClearRACCache
'
'Function Description	 :	Function used to clear cache
'
'Function Parameters	 :   NA
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	Call Fn_Setup_ClearRACCache()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  24-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_ClearRACCache()		
	On Error Resume Next
	'Declaring Variables
	Dim objNetwork,objFSO,objFolder,objSubFolder,objWScriptShell
	Dim sPath
	
	Const DeleteReadOnly = True
	
	'Initially set function return value as False
	Fn_Setup_ClearRACCache = False
	
	'Creating object of Network and File System
	Set objNetwork = CreateObject("WScript.Network")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWScriptShell = CreateObject("WScript.Shell")
	sPath = objWScriptShell.ExpandEnvironmentStrings("%USERPROFILE%")
	
	'Object of User folder
	Set objFolder = objFSO.GetFolder(sPath)
	
	'Object of subfolders and files within User folder
	'Set objSubFolder = objFolder.SubFolders
	
	'Deletes folders and files 
	If objFSO.FolderExists(sPath & "\Teamcenter") Then
		objFSO.DeleteFolder sPath & "\" & "Teamcenter",True
	End If
	
	'Deletes folders and files 
	If objFSO.FolderExists(sPath & "\FCCCache") Then
		objFSO.DeleteFolder sPath & "\" & "FCCCache",True
	End If
	
	'Deletes folders and files 
	If objFSO.FolderExists(sPath & "\Siemens") Then
		objFSO.DeleteFolder sPath & "\" & "Siemens",True
	End If
	
	'Release objects
	Set objNetwork = Nothing
	Set objFSO = Nothing
	Set objFolder = Nothing
	'Set objSubFolder = Nothing
	Set objWScriptShell = Nothing
	
	If Err.Number <> 0 Then
		Fn_Setup_ClearRACCache = False
	Else
		Fn_Setup_ClearRACCache = True
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_GetTestUserDetailsFromExcelOperations
'
'Function Description	 :	Function used to get user details from specified excel file
'
'Function Parameters	 :  1.sAction	: Action Name
'							2.sFilePath	: Excel file name ( Optional )
'							3.sUsername	: Username\Automation ID
'
'Function Return Value	 : 	True or False
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	No
'
'Function Usage		     :	bReturn = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","C:\Ford Mainline\TestData\VSEM210UATTest Users.xlsx","netcom_eng_2")
'Function Usage		     :	bReturn = Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","C:\Ford Mainline\TestData\VSEM210UATTest Users.xlsx","netcom_eng_2")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Vrushali Sahare			    |  24-Feb-2016	    |	 1.0		|		Ganesh Bhosale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_GetTestUserDetailsFromExcelOperations(sAction,sFilePath,ByVal sUsername)
	Err.Clear
	'Declaring Variables
	Dim iCounter
	Dim objFSO,objFile
	Dim sPassword,sGroup,sRole
	Dim sMemberDetails, aMemberDetails
	Dim sTreeNodePath,sUserID
	Dim sContents,aContents,sTempValue
	
	Dim objExcel,objWorkbook
	Dim iloopCount
	
	'Initially set function return value as False
	Fn_Setup_GetTestUserDetailsFromExcelOperations = False
	
	'Get the file path
	If sFilePath = "" Then
		Environment.Value("AutomationDirPath") = Fn_CommonUtil_EnvironmentVariablesOperations("Get","User","AutomationDir","")
		'sFilePath = Environment.Value("AutomationDirPath") & "\TestData\AUT_TestUserDetails.csv"
		If lCase(sAction)="getautomationidbygroupandrole_excel" or lCase(sAction)="getlogindetails_excel" Then
			sFilePath = Environment.Value("AutomationDirPath") & "\TestData\AUT_TestUserDetails.xlsx"
		Else
			sFilePath = Environment.Value("AutomationDirPath") & "\TestData\AUT_TestUserDetails.csv"
		End If
	End If
	
	If lCase(sAction)="getautomationidbygroupandrole_excel" or lCase(sAction)="getlogindetails_excel" Then
		'Creating the Excel object
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = False 
		objExcel.DisplayAlerts = 0

		'Setting the Workbook object
		Set objWorkbook = objExcel.Workbooks.Open(sFilePath)
		iloopCount = 1
	End If
	
	Select Case lCase(sAction)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the login details
		Case "getlogindetails"
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If aContents(0)=sUsername Then
					sUsername=aContents(1)
					sPassword=aContents(2)
					sGroup=aContents(3)
					sRole=aContents(4)
					Fn_Setup_GetTestUserDetailsFromExcelOperations = Cstr(sUsername)& "~" & sPassword & "~" & sGroup & "~" & sRole
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the user name
		Case "getusername"
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If aContents(0)=sUsername Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = aContents(5)
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the user id
		Case "getuserid"
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If aContents(0)=sUsername Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = aContents(1)
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the user id
		Case "getowner"
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If aContents(0)=sUsername Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = aContents(5) & " (" & aContents(1) & ")"					
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the role
		Case "getrole"			
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If aContents(0)=sUsername Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = aContents(4)
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the group
		Case "getgroup"	
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If aContents(0)=sUsername Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = aContents(3)
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the user node hierarchy as displayed in tree node
		Case "getusernodetreepath"	
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			Fn_Setup_GetTestUserDetailsFromExcelOperations = aMemberDetails(2) & "~" & aMemberDetails(3) & "~" & aMemberDetails(0) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sUsername) & ")"			
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the user node hierarchy as displayed in tree node
		Case "getusernodetreepathforprojectselectedmembers"
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			Fn_Setup_GetTestUserDetailsFromExcelOperations = aMemberDetails(2) & "." & aMemberDetails(3) & "~" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sUsername) & ")"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the user node hierarchy as displayed in tree node
		Case "getusernodetreepathforprojectselectionmembers"		
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			sContents = Split(aMemberDetails(2),".")
			For iCounter = UBound(sContents) To 0 Step -1
				If iCounter=UBound(sContents) Then
					aMemberDetails(2)=sContents(iCounter)
				Else
					aMemberDetails(2)=aMemberDetails(2) & "~"  & sContents(iCounter)
				End If
			Next
			Fn_Setup_GetTestUserDetailsFromExcelOperations = aMemberDetails(2) & "~" & aMemberDetails(3) & "~" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername) & " (" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sUsername) & ")"
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		'Case to get Organization tree user node path
		Case "getorganizationtreeusernodepath"							
			sRole=Fn_Setup_GetTestUserDetailsFromExcelOperations("getrole","",sUsername)
			sGroup=Fn_Setup_GetTestUserDetailsFromExcelOperations("getgroup","",sUsername)
			sUserID=Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sUsername)
			
			sGroup=Split(sGroup,".")
			sTreeNodePath=""
			For iCounter=Ubound(sGroup) To 0 Step -1
				If sTreeNodePath<>"" Then
					sTreeNodePath=sTreeNodePath & "~" & sGroup(iCounter)
				Else
					sTreeNodePath=sGroup(iCounter)
				End If
			Next
			sTreeNodePath= sTreeNodePath & "~" & sRole
			sUsername=Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername)			
			sTreeNodePath= sTreeNodePath & "~" & sUsername & " (" & sUserID & ")"
			Fn_Setup_GetTestUserDetailsFromExcelOperations="Organization~" & sTreeNodePath
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Participants tree user node path
		Case "getparticipantstreeusernodepath"							
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			sTempValue=Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			Fn_Setup_GetTestUserDetailsFromExcelOperations ="Participants~" & aMemberDetails(3) & "~" & sTempValue &" ("& Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sUsername) & ")-" & aMemberDetails(2) & "/" & aMemberDetails(3)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get Participants tree user node path
		Case "getparticipantstreeusernodepathext"							
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			sTempValue=Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			Fn_Setup_GetTestUserDetailsFromExcelOperations ="Participants~" & aMemberDetails(3) & " *~" & sTempValue &" ("& Fn_Setup_GetTestUserDetailsFromExcelOperations("getuserid","",sUsername) & ")-" & aMemberDetails(2) & "/" & aMemberDetails(3)		
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		Case "getsdtmember"
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			Fn_Setup_GetTestUserDetailsFromExcelOperations = aMemberDetails(2) & "/" & aMemberDetails(3) & "/" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
		'Case to get Project Teams tree user node path
		Case "getprojectteamsnodepath"
			sMemberDetails = Fn_Setup_GetTestUserDetailsFromExcelOperations("getlogindetails","",sUsername)
			aMemberDetails = Split(sMemberDetails,"~")
			Fn_Setup_GetTestUserDetailsFromExcelOperations =aMemberDetails(3) & "~" & aMemberDetails(2) & "/" & aMemberDetails(3) & "/" & Fn_Setup_GetTestUserDetailsFromExcelOperations("getusername","",sUsername)
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the automation id by group and role
		Case "getautomationidbygroupandrole"		
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			Set objFile = objFSO.OpenTextFile(sFilePath)
			sTempValue=Split(sUsername,"~")
			Do Until objFile.AtEndOfStream
				sContents=objFile.Readline
				aContents=Split(sContents,",")
				If sTempValue(0)=aContents(3) and sTempValue(1)=aContents(4) Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = Cstr(aContents(0))
					Exit do
				End If
			Loop
			objFile.Close
			Set objFSO = Nothing
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the automation id by group and role from xlsx file
		Case "getautomationidbygroupandrole_excel"	
			sTempValue=Split(sUsername,"~")

			Do while not isempty(objExcel.Cells(iloopCount, 1).Value)
				If sTempValue(0)=objExcel.Cells(iloopCount, 4).Value and sTempValue(1)=objExcel.Cells(iloopCount, 5).Value Then
					Fn_Setup_GetTestUserDetailsFromExcelOperations = objExcel.Cells(iloopCount, 1).Value
					Exit do
				End If
				iloopCount = iloopCount+1
			Loop	
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to get the login details
		Case "getlogindetails_excel"
			Do while not isempty(objExcel.Cells(iloopCount, 1).Value)
				If objExcel.Cells(iloopCount, 1).Value = sUsername Then
					sUsername = objExcel.Cells(iloopCount, 2).Value
					sPassword = objExcel.Cells(iloopCount, 3).Value
					sGroup = objExcel.Cells(iloopCount, 4).Value
					sRole = objExcel.Cells(iloopCount, 5).Value
					Fn_Setup_GetTestUserDetailsFromExcelOperations = Cstr(sUsername)& "~" & sPassword & "~" & sGroup & "~" & sRole
					Exit do
				End If
				iloopCount = iloopCount+1
			Loop
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
		'Case to handle invalid request
		Case Else
			Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_Setup_GetTestUserDetailsFromExcelOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ]: No valid case was passed for function [Fn_Setup_GetTestUserDetailsFromExcelOperations]")
	End Select
	
	If lCase(sAction)="getautomationidbygroupandrole_excel" or lCase(sAction)="getlogindetails_excel" Then
		'Close and quit the file
		objExcel.Workbooks.Close
		objExcel.quit

		'Release Objects
		Set objWorkbook = Nothing
		Set objExcel = Nothing
	End If
	
	'Report any unexpected runtime error
	If Err.Number <> 0 Then
		Fn_Setup_GetTestUserDetailsFromExcelOperations = False
		Call Fn_LogUtil_UpdateDetailLog(Environment.Value("TestLogFile"),"<FAIL>:  [ Fn_Setup_GetTestUserDetailsFromExcelOperations ] : Fail to perform operation [ " & Cstr(sAction) & " ] due to error number [ " & Cstr(Err.Number) & " ] with error description [ " & Cstr(Err.Description) & " ]")
	End If
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_SetActionIterationMode
'
'Function Description	 :	Function used to set test case run iteration mode
'
'Function Parameters	 :  1. sIterationMode : Run iteration mode
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	Call Fn_Setup_SetActionIterationMode("oneIteration")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  10-Mar-2016	    |	 1.0		|		Kundan Kudale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public function Fn_Setup_SetActionIterationMode(sIterationMode)
	On Error Resume Next
	'Declaring varaibles
	Dim objQuickTestApp
	If sIterationMode="" Then
		sIterationMode="oneIteration"
	End If
	'Creating object of QuickTest Application
	Set objQuickTestApp = CreateObject("QuickTest.Application")
	'Setting run iteration mode
	objQuickTestApp.Test.Settings.Run.IterationMode = sMode
    objQuickTestApp.Options.Run.RunMode = "Normal"
	'Releasing object of QuickTest Application
	Set objQuickTestApp = Nothing	
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_GenerateObjectInformation
'
'Function Description	 :	Function is used to generate information of application objects
'
'Function Parameters	 :  1. sAction : Action name 
'							2. sValue  : Value
'
'Function Return Value	 : 	False\Name\ID
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_Setup_GenerateObjectInformation("getname","Item")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  08-Jul-2016	    |	 1.0		|		Kundan Kudale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_GenerateObjectInformation(sAction,ByVal sValue)
	'Initially function return false
	Fn_Setup_GenerateObjectInformation=False
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "getname"
			Select Case sValue
				Case "Item","Part","Design","Dataset"
					sValue=Replace(sValue," ","")
					Fn_Setup_GenerateObjectInformation="AUT_" & Cstr(sValue & Cstr(Fn_CommonUtil_GenerateRandomNumber(5)))
				Case Else
					sValue=Replace(sValue," ","")
					Fn_Setup_GenerateObjectInformation="AUT_" & Cstr(sValue & Cstr(Fn_CommonUtil_GenerateRandomNumber(5)))
			End Select				
	End Select
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header Start
'Function Name			 :	Fn_Setup_DeleteCacheFiles
'
'Function Description	 :	Function used to delete cache files
'
'Function Parameters	 :  NA
'
'Function Return Value	 : 	NA
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :  bReturn = Fn_Setup_DeleteCacheFiles()
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	 Reviewer			|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  16-Mar-2016	    |	 1.0		|	Kundan Kudale	 	| 	Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Function Fn_Setup_DeleteCacheFiles()
		Reporter.Filter = rfDisableAll
		On Error Resume Next
		'Variable Declaration
		Dim objFolder,objSubFolder,objWscriptShell,objFSO,objWscriptNetwork,objFiles
		Dim sFolderName
		Dim aFolderName
		Dim iCounter
		
		'Constant Variable Declaration
		Const DeleteReadOnly = True
		'Creating Object of Shell
		Set objWscriptShell = CreateObject("Wscript.Shell")
		'Creating Object of Filesystem
		Set objFSO = CreateObject("Scripting.FilesystemObject") 
		'Creating Object of Network
		Set objWscriptNetwork = CreateObject("WScript.Network")		
		Set objFolder = objFSO.GetFolder("C:\Documents and Settings\" & objWscriptNetwork.UserName)	
		
		sFolderName = "FCCCache"
		'Deleting FCCCache folders
		Set objSubFolder = objFolder.SubFolders
		For Each iCounter in objSubFolder
			aFolderName = split(iCounter.Name, "_",-1,1)
			if sFolderName = aFolderName(0) Then
			  objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True 		
			End if 
		Next
		'Deleting Teamcenter folders
		If objFSO.FolderExists(objFolder & "\Teamcenter") Then			
			objFSO.DeleteFolder objFolder & "\" & "Teamcenter",True	
		End If
		'Deleting Siemens folders	
		If objFSO.FolderExists(objFolder & "\Siemens") Then
			objFSO.DeleteFolder objFolder & "\" & "Siemens",True 
		End If
		If objFSO.FolderExists(objFolder & "\.TcIC") Then
			objFSO.DeleteFolder objFolder & "\" & ".TcIC",True 	
		End If
		If objFSO.FolderExists(objFolder & "\.swt") Then
			objFSO.DeleteFolder objFolder & "\" & ".swt",True 
		End If
		sFolderName = ".Administrator"
		
		For Each iCounter in objSubFolder
			aFolderName = split(iCounter.Name, "_",-1,1)
			If sFolderName = aFolderName(0) then
				'Deleting .Administrator folders
				objFSO.DeleteFolder objFolder & "\" & iCounter.Name,True		
			End if 
		Next
		
		For Each iCounter in objSubFolder
			aFolderName = split(iCounter.Name, "_",-1,1)
			'Check the existing of .Administrator and Teamcenter folders	
			If iCounter.Name = "Teamcenter" OR aFolderName(0)= ".Administrator" then	
				Call Fn_Setup_DeleteCacheFiles()
			End if
		Next			
		
		'Deleting Fcc Files
		objFSO.DeleteFile (objFolder & "\" & "fcc.*"),DeleteReadOnly
		
		If objFSO.FileExists(objFolder & "\TCPLM-JAVA.txt") Then
			'Deleting Fcc Files
			objFSO.DeleteFile (objFolder & "\" & "TCPLM-JAVA.txt"),DeleteReadOnly
		End If
			
		'Releasing all objects
		Set objWscriptNetwork = Nothing
		Set objWscriptShell = Nothing
		Set objSubFolder = Nothing
		Set objFolder = Nothing 
		Set objFSO = Nothing
		
		Reporter.Filter = rfEnableAll
		Err.Clear
End Function
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - @Function Header Start 
'Function Name			 :	Fn_Setup_GetNXTestData
'
'Function Description	 :	Function is used to get information of NX test data Details
'
'Function Parameters	 :  1. sAction : Action name 
'							2. sValue  : Value
'
'Function Return Value	 : 	False\value
'
'Wrapper Function	     : 	NA
'
'Function Pre-requisite	 :	NA
'
'Function Usage		     :	bReturn=Fn_Setup_GetNXTestData("GetModelTemplateFileName","")
'                       
'History			     :
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Developer Name				|	Date			|	Rev. No.   	|	      Reviewer		|	Changes Done	
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'	Sandeep Navghane		    |  30-Oct-2017	    |	 1.0		|		Kundan Kudale	| 		Created
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -  - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -@Function Header End
Public Function Fn_Setup_GetNXTestData(sAction,ByVal sValue)
	'Declare variables
	
	Select Case sAction
		' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
		Case "GetModelTemplateFileName"
			
	End Select
End Function