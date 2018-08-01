Call Fn_StartCommand()

''----------------------------------------------------------------------------------------------------------------------
''	Developer Name
''----------------------------------------------------------------------------------------------------------------------
'' 	Sandeep
''----------------------------------------------------------------------------------------------------------------------
Public Function Fn_StartCommand_XMLNodeValueOperations(sAction,sNodeName,sValue)
	
	Fn_StartCommand_XMLNodeValueOperations=False
	
	Dim objXMLDoc
	Dim objChildNodes
	Dim objSelectNode
	Dim intNodeLength
	Dim intNodeCount
	Dim intChildNodeCount
	Dim strNodeSting
	Dim objSelectNodeName
	Dim objShell
	Dim objUsrEnv
	Dim UserVar
	
	UserVar=""
	
	Set objShell = CreateObject("WScript.Shell")
	Set objUsrEnv = objShell.Environment("User")
	UserVar = objUsrEnv("AutomationDir")
    Set objUsrEnv = Nothing
	Set objShell = Nothing
	
	set objXMLDoc=CreateObject("Microsoft.XMLDOM")												' Create XMLDOM object
	objXMLDoc.async="false"
	objXMLDoc.load(UserVar & "\Utilities\FunctionUtilityVBS\CATIA_CommandInformation.xml")																	' Loading QTP Environment XML

	If (objXMLDoc.parseError.errorCode <> 0) Then
		Fn_StartCommand_XMLNodeValueOperations = False
	Else
		intNodeLength = objXMLDoc.getElementsByTagName("Variable").length
		For intNodeCount = 0 to (intNodeLength - 1)
			Set objSelectNodeName = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Name")
				
			If Trim(objSelectNodeName.Text)=Trim(sNodeName) Then
				Set objSelectNode = objXMLDoc.SelectSingleNode("/Environment/Variable[" & intNodeCount &"]/Value")
				If sAction="get" Then
					Fn_StartCommand_XMLNodeValueOperations = objSelectNode.Text
				ElseIf sAction="set" Then
					objSelectNode.Text=sValue
					objXMLDoc.Save(UserVar & "\Utilities\FunctionUtilityVBS\CATIA_CommandInformation.xml")
					Fn_StartCommand_XMLNodeValueOperations=True
				End If
				Exit For
			End If
		Next
		Set objSelectNode = nothing 
		Set objChildNodes = nothing
		Set objXMLDoc = nothing

		If Fn_StartCommand_XMLNodeValueOperations = "" Then
			Fn_StartCommand_XMLNodeValueOperations = False
		End If
	
	End if	

End Function
''----------------------------------------------------------------------------------------------------------------------
''	Developer Name
''----------------------------------------------------------------------------------------------------------------------
'' 	Sandeep
''----------------------------------------------------------------------------------------------------------------------
Function Fn_StartCommand()
	Dim CATIA
	Dim sCommand
	Call Fn_StartCommand_XMLNodeValueOperations("set","Result","False")
	
	sCommand=Fn_StartCommand_XMLNodeValueOperations("get","CommandName","")
	If sCommand=False Then
		Exit Function
	End If
	
	Call Fn_StartCommand_XMLNodeValueOperations("set","Result","True")
	
	Set CATIA=GetObject("","CATIA.Application")
	CATIA.StartCommand sCommand
	Set CATIA=Nothing
	
	If Err.Number<0 Then
		Call Fn_StartCommand_XMLNodeValueOperations("set","Result","False")
	Else
		Call Fn_StartCommand_XMLNodeValueOperations("set","Result","True")
	End If
End Function