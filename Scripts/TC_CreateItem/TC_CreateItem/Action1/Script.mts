
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'            Module Name             :   EBOM
'
'            Testcase Name           :   Create & Modify GET Document in Teamcenter
'
'            Test Objective          :   Create & Modify GET Document in Teamcenter
'
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'            Developer Name          |       Date                |   Teamcenter Release      |  Reviewer
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'            Pallavi Jadhav          |       11-January-2017       |   Teamcenter 11.2          |   Pallavi Jadhav
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
Option Explicit
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Variable Declaration
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Initialize Testcase
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
LoadAndRunAction "FMW_Setup\FMW_Setup_TestcaseInit","FMW_Setup_TestcaseInit", oneIteration
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 1
'Description     : Login to Teamcenter with any of the following Roles:Engineer,Designer, Engineering Manager,Material Technical Advisor  Note: Only above roles can create a GET Document Itemt.
'Expected result : Logged into Teamcenter Successfully..
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="Login to Teamcenter with any of the following Roles:Engineer,Designer, Engineering Manager,Material Technical Advisor  Note: Only above roles can create a GET Document Itemt."
GBL_STEP_EXPECTED_RESULT="Logged into Teamcenter Successfully.."
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)

'Launch Teamcenter
LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_ReuseTcSession","RAC_LoginUtil_ReuseTcSession",OneIteration,True,True,"",DataTable.Value("AutomationUserID","Global"),"","portalbat",""
'Modify Group Role
'LoadAndRunAction "RAC_Common\RAC_Common_UserSessionSettingsOperations","RAC_Common_UserSessionSettingsOperations",OneIteration,"ModifySession",DataTable.Value("AutomationUserID","Global"),"",""

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 2
'Description     : In the My Teamcenter perspective, Select the desired folder to create new GET Document.Create Document Item using File-> New-> Item Menu.
'Expected result : New part dialog should be displayed
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="In the My Teamcenter perspective, Select the desired folder to create new GET Document.Create Document Item using File-> New-> Item Menu."
GBL_STEP_EXPECTED_RESULT="New part dialog should be displayed"
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)

' Set My Teamcenter Perspective.
LoadAndRunAction "RAC_Common\RAC_Common_SetResetPerspective","RAC_Common_SetResetPerspective",OneIteration,GBL_PERSPECTIVE_MYTEAMCENTER,True,True

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 3
'Description     : Select "GET Document" from the list. Click on Next Button.
'Expected result : GET Document Item details window should be displayed.
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="Select GET Document from the list. Click on Next Button."
GBL_STEP_EXPECTED_RESULT="GET Document Item details window should be displayed."
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 4
'Description     : User enters Item id ( Acquired from Item ID tool from Control Card) in the ID field. Enter all the mandatory values : Name, IP Classifciation etc
'Expected result : All mandatory fields should be populated.
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="User enters Item id ( Acquired from Item ID tool from Control Card) in the ID field. Enter all the mandatory values : Name, IP Classifciation etc"
GBL_STEP_EXPECTED_RESULT="All mandatory fields should be populated."
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 5
'Description     : Click on Finish button to complete creation of GET Document item.
'Expected result : GET Document item should be created in selected folder.
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="Click on Finish button to complete creation of GET Document item."
GBL_STEP_EXPECTED_RESULT="GET Document item should be created in selected folder."
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)

'Create Test case Folder
LoadAndRunAction "RAC_Common\RAC_Common_CreateTestCaseFolder","RAC_Common_CreateTestCaseFolder",OneIteration,"","",""
' Expand Test Case Folder 
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Expand", GBL_TESTCASE_FOLDER_PATH,""
'Select  Test Case Folder 
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "Select", GBL_TESTCASE_FOLDER_PATH,""
'Create GET Document
LoadAndRunAction "RAC_Common\RAC_Common_CreateItem","RAC_Common_CreateItem", OneIteration, "autobasiccreate", Datatable.Value("ItemType"), "menu", "myteamcenter", ""

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 6
'Description     : Search for the GET Document created in previous test scipt.
'Expected result : GET Document selected and opened in new window
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="Search for the GET Document created in previous test scipt."
GBL_STEP_EXPECTED_RESULT="GET Document selected and opened in new window"
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)

sItemNode = GBL_TESTCASE_FOLDER_PATH & "~" & DataTable.Value("RACItemID","Global") &"/"&RACItemRevisionID

'Verify existance of GET Document
LoadAndRunAction "RAC_MyTeamcenter\RAC_MyTc_NavigationTreeOperations", "RAC_MyTc_NavigationTreeOperations", oneIteration, "VerifyExist",sItemNode ,""

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Step No.        : Step 7
'Description     : Logout from Teamcenter
'Expected result : Logged out successfully.
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
GBL_STEP_DESCRIPTION="Logout from Teamcenter"
GBL_STEP_EXPECTED_RESULT="Logged out successfully."
Call Fn_LogUtil_PrintStepHeaderLog(GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT)
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Exit From Test Case
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
LoadAndRunAction"FMW_Setup\FMW_Setup_TestcaseExit","FMW_Setup_TestcaseExit", oneIteration,"RAC",True

Call Fn_ExitTest()

Function Fn_ExitTest()
 ExitTest
End Function

