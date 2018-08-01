'! @Name 		RAC_MPP_SetIndexOfMPPAppletFromTab
'! @Details 	To set index of Applet with the help of Tab Names
'! @Author 		Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 	Kundan Kudale kundan.kudale@sqs.com
'! @Date 		08 Jan 2017
'! @Version 	1.0
'! @Example 	LoadAndRunAction "RAC_ManufacturingProcessPlanner\RAC_MPP_SetIndexOfMPPAppletFromTab","RAC_MPP_SetIndexOfMPPAppletFromTab",OneIteration

Option Explicit
Err.Clear

'Declaring variables
Dim sTabID, sTabRevID, sTabName,sTableID, sTableRevID,sTableName,sTableTopNode
Dim objMPPApplet,objManufacturingProcessPlanner
Dim iTabIndex,iAppletCount
Dim aTopNode, aTabText

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

sTabID = ""
sTabRevID = ""
sTabName = ""

Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordName","","")
Call Fn_CommonUtil_DataTableOperations("AddColumn","ReusableActionWordReturnValue","","")
DataTable.SetCurrentRow 1		
DataTable.Value("ReusableActionWordName","Global")= "RAC_MPP_SetIndexOfMPPAppletFromTab"
DataTable.Value("ReusableActionWordReturnValue","Global")= "False"

'Creating object of [ MPPApplet ] applet
Set objMPPApplet=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ManufacturingProcessPlanner_OR","wjapt_MPPApplet","")
'Creating object of [ Manufacturing Process Planner ] window
Set objManufacturingProcessPlanner=Fn_FSOUtil_XMLFileOperations("getobject","RAC_ManufacturingProcessPlanner_OR","jwnd_ManufacturingProcessPlanner","")

If objManufacturingProcessPlanner.JavaTab("jtab_ViewAll").Exist(3) = True Then
	If lcase(objManufacturingProcessPlanner.JavaTab("jtab_ViewAll").GetROProperty("value")) <> "base view" Then
		sTabName = lcase(objManufacturingProcessPlanner.JavaTab("jtab_ViewAll").GetROProperty("value"))
		For iAppletCount = 0 to 10
			objMPPApplet.SetTOProperty "Index",iAppletCount
			If objMPPApplet.JavaTable("CMEBOMTreeTable").Exist(1) Then
				If lcase(sTabName) = lcase(objMPPApplet.JavaTable("jtbl_CMEBOMTreeTable").Object.getValueAt(0,0).toString()) Then
					DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
					'Releasing objects
					Set objMPPApplet=Nothing
					Set objManufacturingProcessPlanner=Nothing
					ExitAction
				End If
			End If
		Next
	End If
End If	

'Getting Selected Tab Index
iTabIndex = cInt(objManufacturingProcessPlanner.JavaObject("jobj_RACTabFolderWidget").Object.getSelectedTabIndex)
sTabName = objManufacturingProcessPlanner.JavaObject("jobj_RACTabFolderWidget").Object.getItem(iTabIndex ).text

If instr(sTabName,";") > 0  Then
	sTabName = replace(sTabName,"/","-")
	sTabName = replace(sTabName,";1-","-")
	aTabText = split(sTabName, "-")
	sTabID = aTabText(0)
	sTabRevID = aTabText(1)
	sTabName = aTabText(2)
ElseIf instr(sTabName,"-") > 0  Then
	aTabText = split(sTabName, "-")
	sTabID = aTabText(0)
	sTabName = aTabText(1)
Else
	'Name
	sTabName = sTabName
End If

For iAppletCount = 0 to 10
	objMPPApplet.SetTOProperty "Index",iAppletCount	
	If objMPPApplet.JavaTable("jtbl_CMEBOMTreeTable").Exist(2) Then
		sTableTopNode = objMPPApplet.JavaTable("jtbl_CMEBOMTreeTable").Object.getValueAt(0,0).toString()
		sTableID = ""
		sTableRevID = ""
		sTableName = ""
		If instr(sTableTopNode,";") > 0  Then
			'id rev name
			sTableTopNode = replace(sTableTopNode,"/","-")
			sTableTopNode = replace(sTableTopNode,";1-","-")
			sTableTopNode = trim(replace(sTableTopNode,"(View)",""))
			aTopNode = split(sTableTopNode, "-")
			sTableID = aTopNode(0)
			sTableRevID = aTopNode(1)
			sTableName = aTopNode(2)
		ElseIf instr(sTableTopNode,"-") > 0  Then
			'id name
			aTopNode = split(sTableTopNode, "-")
			sTableID = aTopNode(0)
			sTableName = aTopNode(1)
		Else
			' Name
			sTableName = sTableTopNode
		End If

		If sTabRevID  <> "" Then		'Tab Revision ID does not exist when we send a Struc Context, therefore below change in code.... Amit T - 04 - July - 2012
			If sTabRevID = sTableRevID Then
				If sTabID <> "" Then
					'match Name and ID
					If sTabID = sTableID AND sTabName = sTableName Then
						DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
						Exit For
					End If
				ElseIf sTabName = sTableName Then
					DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
					Exit For
				End If
			End If
		Else
			If sTabID <> "" Then
				'match Name and ID
				If sTabID = sTableID AND sTabName = sTableName Then
					DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
					Exit For
				End If
			Else
				DataTable.Value("ReusableActionWordReturnValue","Global")= "True"
					Exit For
			End If
		End If
	End If
Next

'Releasing objects
Set objMPPApplet=Nothing
Set objManufacturingProcessPlanner=Nothing
