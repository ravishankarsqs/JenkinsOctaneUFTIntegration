' This test was created using HP ALM
'! @Name 		RAC_LoginUtil_CloseAllDialogs
'! @Details 	To close all open dialogs in Teamcenter application
'! @Author 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 	Kundan Kudale kundan.kudale@sqs.com
'! @Date 		25 Mar 2016
'! @Version 	1.0
'! @Example 	LoadAndRunAction "RAC_LoginUtil\RAC_LoginUtil_CloseAllDialogs","RAC_LoginUtil_CloseAllDialogs",oneIteration

Option Explicit
Err.Clear

'Declaring variables
Dim iCounter,iCount
Dim objTcDefaultApplet,objDefaultWindow,objJavaApplet,objJavaDialog,objJavaWindow,objChild

'Creating object of [ teamcenter Default ] window
Set objDefaultWindow =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_DefaultWindow","")
'Creating object of [ TcDefaultApplet ] window
Set objTcDefaultApplet =Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","jwnd_TcDefaultApplet","")
'Creating object of [ teamcenter main Java ] applet
Set objJavaApplet = Fn_FSOUtil_XMLFileOperations("getobject","RAC_LoginUtil_OR","japt_JavaApplet","")
	
'Checking existence of teamcenter default window
If Fn_UI_Object_Operations("Fn_RACLoginUtil_CloseAllDialogs","Exist", objDefaultWindow, "","","")=False Then
	ExitAction
End If

'Closing all open dialogs
Call Fn_Setup_ReporterFilter("DisableAll")

For iCounter=1 to 3
	'Close All Java dialogs appears under TcDefaultApplet object
	Set objJavaDialog=Description.Create()
	objJavaDialog("Class Name").Value="JavaDialog"
	If objTcDefaultApplet.ChildObjects(objJavaDialog).Count <> 0 Then
		Set objChild=objTcDefaultApplet.ChildObjects(objJavaDialog)
		For iCount=0 to objChild.count-1
			objChild(iCount).Close
			wait GBL_MIN_MICRO_TIMEOUT
		Next
		Set objChild=Nothing	
	End If
	Set objJavaDialog=Nothing
	
	'Close All Java dialogs appears under teamcenter main Java Applet object
'	Set objJavaDialog=Description.Create()
'	objJavaDialog("Class Name").Value="JavaDialog"	
'	If objJavaApplet.ChildObjects(objJavaDialog).Count <> 0 Then
'		Set objChild=objJavaApplet.ChildObjects(objJavaDialog)
'		For iCount=0 to objChild.count-1
'			objChild(iCount).Close
'			wait GBL_MICRO_TIMEOUT
'		Next
'		Set objChild=Nothing
'	End If
'	Set objJavaDialog=Nothing
			
	'Close All Java dialogs appears under teamcenter Default Window
	Set objJavaDialog=Description.Create()
	objJavaDialog("Class Name").Value="JavaDialog"	
	If objDefaultWindow.ChildObjects(objJavaDialog).Count <> 0 Then
		Set objChild=objDefaultWindow.ChildObjects(objJavaDialog)
		For iCount=0 to objChild.count-1
			objChild(iCount).Close
			wait GBL_MIN_MICRO_TIMEOUT
		Next
		Set objChild=Nothing		
	End If
	Set objJavaDialog=Nothing
	
	'Close All Java windows appears under teamcenter Default Window
	Set objJavaWindow=Description.Create()
	objJavaWindow("Class Name").Value="JavaWindow"
	If objDefaultWindow.ChildObjects(objJavaWindow).Count <> 0 Then
		Set objChild=objDefaultWindow.ChildObjects(objJavaWindow)
		For iCount=0 to objChild.count-1
			If objChild(iCount).GetROProperty("toolkit class")="sun.awt.windows.WEmbeddedFrame" and objChild(iCount).GetROProperty("title")="" Then
				'do nothing
			Else
				objChild(iCount).Close
				wait GBL_MIN_MICRO_TIMEOUT
			End If
		Next
		Set objChild=Nothing		
	End If
	Set objJavaWindow=Nothing	
Next
Call Fn_Setup_ReporterFilter("EnableAll")

'Releasing all created objects
Set objDefaultWindow = Nothing
Set objTcDefaultApplet = Nothing
Set objJavaApplet = Nothing

