Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for time in seconds
Public GBL_DEFAULT_TIMEOUT, GBL_MICRO_TIMEOUT, GBL_MIN_TIMEOUT, GBL_MAX_TIMEOUT,GBL_DEFAULT_MIN_TIMEOUT,GBL_MIN_MICRO_TIMEOUT,GBL_ZERO_TIMEOUT
GBL_MAX_TIMEOUT = 150 'time in seconds
GBL_DEFAULT_TIMEOUT = 20 'time in seconds
GBL_DEFAULT_MIN_TIMEOUT = 10 'time in seconds
GBL_MIN_TIMEOUT = 5 'time in seconds
GBL_MIN_MICRO_TIMEOUT = 2 'time in seconds
GBL_MICRO_TIMEOUT = 1 'time in seconds
GBL_ZERO_TIMEOUT = 0 'time in seconds
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for number of sync iterations
Public GBL_DEFAULT_SYNC_ITERATIONS, GBL_MICRO_SYNC_ITERATIONS, GBL_MIN_SYNC_ITERATIONS, GBL_MAX_SYNC_ITERATIONS,GBL_MIN_MICRO_SYNC_ITERATIONS,GBL_APP_SYNC_ITERATIONS
GBL_MAX_SYNC_ITERATIONS = 10 'number of iterations
GBL_DEFAULT_SYNC_ITERATIONS = 5 'number of iterations
GBL_MIN_SYNC_ITERATIONS = 3 'number of iterations
GBL_MIN_MICRO_SYNC_ITERATIONS = 2 'number of iterations
GBL_MICRO_SYNC_ITERATIONS = 1 'number of iterations
GBL_APP_SYNC_ITERATIONS=""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for test case log type
Public GBL_LOG_STEP_HEADER, GBL_LOG_FAIL_VERIFICATION, GBL_LOG_PASS_VERIFICATION,GBL_LOG_FAIL_ACTION, GBL_LOG_PASS_ACTION
GBL_LOG_STEP_HEADER = "step_header"
GBL_LOG_FAIL_VERIFICATION = "fail_verification"
GBL_LOG_PASS_VERIFICATION = "pass_verification"
GBL_LOG_FAIL_ACTION = "fail_action"
GBL_LOG_PASS_ACTION = "pass_action"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for test case step information
Public GBL_STEP_NUMBER,GBL_STEP_DESCRIPTION,GBL_STEP_EXPECTED_RESULT,GBL_STEP_EXECUTION_TIME
GBL_STEP_NUMBER = 0

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for application names
Public GBL_APP_NAME_TO_EXIT_ON_FAILURE,GBL_APP_NAME_TO_SYNC,GBL_CURRENT_EXECUTABLE_APP

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for execution times
Public GBL_FUNCTION_EXECUTION_START_TIME,GBL_FUNCTION_EXECUTION_END_TIME
Public GBL_TEST_EXECUTION_START_TIME, GBL_TEST_EXECUTION_END_TIME,GBL_TEST_EXECUTION_TOTAL_TIME
GBL_TEST_EXECUTION_START_TIME=""
GBL_TEST_EXECUTION_END_TIME=""
GBL_TEST_EXECUTION_TOTAL_TIME=""
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for automation standard prefix
Public GBL_AUTOMATION_PREFIX
GBL_AUTOMATION_PREFIX="SQS_AUT"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for folder path of navigation tree
Public GBL_AUTOMATEDTEST_FOLDER_PATH,GBL_NEWSTUFF_FOLDER_PATH,GBL_TESTCASE_FOLDER_PATH
GBL_AUTOMATEDTEST_FOLDER_PATH="Home~AutomatedTest"
GBL_NEWSTUFF_FOLDER_PATH="Home~Newstuff"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for counters
Public GBL_VERIFICATION_COUNTER

GBL_VERIFICATION_COUNTER=1
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for datatable current set row number
Public GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for myworklist tree node suffix
Public GBL_MW_TN_TASKSTOPERFORM,GBL_MW_TN_PERFORMSIGNOFFS,GBL_MW_TN_TASKSTOTRACK
GBL_MW_TN_TASKSTOPERFORM = "Tasks To Perform"
GBL_MW_TN_PERFORMSIGNOFFS = "(perform-signoffs)"
GBL_MW_TN_TASKSTOTRACK = "Tasks To Track"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for Teamcenter Perspective Names
Public GBL_PERSPECTIVE_MYTEAMCENTER,GBL_PERSPECTIVE_STRUCTUREMANAGER,GBL_PERSPECTIVE_SYSTEMENGINERRING,GBL_PERSPECTIVE_REQUIREMENTMANAGER
Public GBL_PERSPECTIVE_CHANGEMANAGER,GBL_PERSPECTIVE_PROJECT
GBL_PERSPECTIVE_MYTEAMCENTER = "My Teamcenter"
GBL_PERSPECTIVE_STRUCTUREMANAGER = "Structure Manager"
GBL_PERSPECTIVE_SYSTEMENGINERRING = "System Engineering"
GBL_PERSPECTIVE_REQUIREMENTMANAGER = "Requirement Manager"
GBL_PERSPECTIVE_CHANGEMANAGER = "Change Manager"
GBL_PERSPECTIVE_PROJECT = "Project"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for function log
Public GBL_FUNCTIONLOG
GBL_FUNCTIONLOG=""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables for log updation time
Public GBL_LAST_LOG_UPDATION_TIME

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables used for function Fn_RAC_ReadyStatusSync
Public GBL_TCOBJECTS_SYNC_FLAG,GBL_TCOBJECTS_SYNC_XAXIS,GBL_TCOBJECTS_SYNC_YAXIS
GBL_TCOBJECTS_SYNC_FLAG = False

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables used to store java tree node information
Public GBL_JAVATREE_NODEBOUNDS_OBJECT,GBL_JAVATREE_CURRENTNODE_OBJECT

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to indicate whether to disable print and update log functionality
Public GBL_DISABLEPRINTANDUPDATELOG_REPORTING
GBL_DISABLEPRINTANDUPDATELOG_REPORTING=False

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store additional log information
Public GBL_LOG_ADDITIONAL_INFORMATION,GBL_CATIA_PARTNAME
GBL_LOG_ADDITIONAL_INFORMATION = ""
GBL_CATIA_PARTNAME = ""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables to store teamcenter syslog image path
Public GBL_SYSLOG_IMAGE_PATH
GBL_SYSLOG_IMAGE_PATH=""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variables to store application Opened in test case
Public GBL_APPLICATIONS_OPENED_IN_TEST
GBL_APPLICATIONS_OPENED_IN_TEST="NA"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store teamcenter application invoke option
Public GBL_TEAMCENTER_INVOKE_OPTION
GBL_TEAMCENTER_INVOKE_OPTION=""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store user id logged into teamcenter application
Public GBL_TEAMCENTER_LAST_LOGGEDIN_USERID,GBL_CATIA_TEAMCENTE_SAVE_OPTION_LOGIN_FLAG,GBL_CATIA_TEAMCENTER_INTEGRATION_LOGIN_FLAG,GBL_LOADINCATIAPROCESS_FLAG
GBL_TEAMCENTER_LAST_LOGGEDIN_USERID=""
GBL_CATIA_TEAMCENTE_SAVE_OPTION_LOGIN_FLAG=False
GBL_LOADINCATIAPROCESS_FLAG=False
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store HP UFT product name
Public GBL_HP_QTP_PRODUCTNAME
GBL_HP_QTP_PRODUCTNAME="HP Unified Functional Testing"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store Teamcenter login type
Public GBL_TC_LOGINTYPE
GBL_TC_LOGINTYPE="NA"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store TPDM login user details
Public GBL_TPDM_LOGGEDIN_USER
GBL_TPDM_LOGGEDIN_USER = ""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store catia enviroenment name to launch from custom launch case
Public GBL_CATIA_ENVIRONMENT_NAME_FOR_CUSTOM_LAUNCH,GBL_WRAPPER_CATIA_APP_LAUNCH_FLAG,GBL_WRAPPER_NX_APP_LAUNCH_FLAG
GBL_CATIA_ENVIRONMENT_NAME_FOR_CUSTOM_LAUNCH = "NA"
GBL_WRAPPER_CATIA_APP_LAUNCH_FLAG=False
GBL_WRAPPER_NX_APP_LAUNCH_FLAG=False
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store flag of test case exit from FMW_Setup_TestcaseExit action word : Flag value gets change to True in FMW_Setup_TestcaseExit action word
Public GBL_TESTCASE_EXIT_FLAG
GBL_TESTCASE_EXIT_FLAG = False

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store flag of node\item\object send to structure manager from nav tree
Public GBL_SENDTOPSM_FROM_NAVTREE_FLAG,GBL_SENDTOPSM_FROM_NAVTREE_NODEPATH
GBL_SENDTOPSM_FROM_NAVTREE_FLAG = False
GBL_SENDTOPSM_FROM_NAVTREE_NODEPATH = ""

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store ALM related information
Public GBL_TESTCASE_ID,GBL_TESTSET_NAME,GBL_LASTEXECUTED_ACTIONWORD_NAME,GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME
GBL_TESTCASE_ID = "NA"
GBL_TESTSET_NAME = "NA"
GBL_LASTEXECUTED_ACTIONWORD_NAME = "NA"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= "NA"

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
'Global variable to store teamcenter enviroenment name to launch from custom launch case
Public GBL_Tc_ENVIRONMENT_NAME_FOR_CUSTOM_LAUNCH
GBL_Tc_ENVIRONMENT_NAME_FOR_CUSTOM_LAUNCH="NA"