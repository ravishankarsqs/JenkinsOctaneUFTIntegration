'! @Name 			RAC_Common_CreateItem
'! @Details 		Action word to perform operations on New Item creation dialog. eg. Item basic create , item detail create
'! @InputParam1 	sAction 			: String to indicate what action is to be performed
'! @InputParam2 	sItemType 			: Item Type
'! @InputParam3 	sInvokeOption		: New Item creation dialog invoke option
'! @InputParam4 	sPerspective	 	: Perspective name in which user wants to perform operations on New Item dialog
'! @InputParam5 	sButton 			: Button Name
'! @Author 			Sandeep Navghane sandeep.navghane@sqs.com
'! @Reviewer 		Kundan Kudale kundan.kudale@sqs.com
'! @Date 			07 Jul 2016
'! @Version 		1.0
'! @Example 		LoadAndRunAction "RAC_Common\RAC_Common_CreateItem","RAC_Common_CreateItem",OneIteration,"autobasiccreate","ItemType_Item","menu","myteamcenter",""

Option Explicit
Err.Clear

'Declaring variables
Dim sAction,sItemType,sInvokeOption,sPerspective,sButton
Dim sItemRevisionID,sItemID,sCurrentItemName,sTempItemType,sName, sNode,sLOVTreeNode
Dim objNewItem,objWScriptShell,objLOVTree
Dim iItemCount,iCounter,iItemNodeCount, iNameCount
Dim bFlag
Dim aName,aLOVTreeNode
Dim aProperty,aValues,sTempPropertyName,sTempName,aFieldNames
Dim sFieldNames

'Assigning currently executable application name
GBL_CURRENT_EXECUTABLE_APP="RAC"

'Reading all paramenter values
sAction = Parameter("sAction")
sItemType = Parameter("sItemType")
sInvokeOption = Parameter("sInvokeOption")
sPerspective = Parameter("sPerspective")
sButton = Parameter("sButton")

sTempItemType = sItemType

GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER=DataTable.GlobalSheet.GetCurrentRow

'Invoking [ New Item ] dialog
Select Case LCase(sInvokeOption)
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "menu",""
		LoadAndRunAction "RAC_Common\RAC_Common_MenuOperations","RAC_Common_MenuOperations",OneIteration,"Select","FileNewItem"
	'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
	Case "nooption"
		'use this case when user whants to invoke dilaog from outside of function
End Select

If sPerspective="" Then
	sPerspective=Fn_RAC_GetActivePerspectiveName("")
End If

'Creating object of [ New Item ] dialog
Select Case sPerspective
	Case "myteamcenter","","structuremanager"
		'Creating object of [ New Item ]
		Set objNewItem=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jwnd_NewItem","")
		'Creating object of LOV tree
		Set objLOVTree=Fn_FSOUtil_XMLFileOperations("getobject","RAC_Common_FU_OR","jtee_LOVTree","")
End Select

GBL_LASTEXECUTED_ACTIONWORD_NAME = "RAC_Common_CreateItem"
GBL_LASTEXECUTED_ACTIONWORD_CASE_NAME= sAction

'Checking existance of new item dialog
If Fn_UI_Object_Operations("RAC_Common_CreateItem","Exist", objNewItem, GBL_DEFAULT_TIMEOUT,"","")=False Then
	Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to perform operation [ " & Cstr(sAction) & " ] on [ New Item ] creation dialog as dialog does not exist","","","","","")
	Call Fn_ExitTest()
End If

'Setting Item count
If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" Then
	DataTable.SetCurrentRow 1
	Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemCount","","")
	iItemCount=Fn_CommonUtil_DataTableOperations("GetValue","RACItemCount","","")
	If iItemCount="" Then
		iItemCount=1
	Else
		iItemCount=iItemCount+1
	End If
	Call Fn_CommonUtil_DataTableOperations("SetValue","RACItemCount",iItemCount,"")
End If

'Get actual item type name
If sItemType<>"" Then
	sItemType=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",sItemType,""))
	sTempItemType=sItemType
End If

'Capturing execution start time
Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureStartTime","RAC_Common_CreateItem",sAction,"","")


If LCase(sInvokeOption)<>"verifyitemtypenotexist" Then
	If Lcase(sAction)= "autobasiccreate" or Lcase(sAction)="basiccreate" or Lcase(sAction)="verifylovtreenode" or lcase(sAction)="verifylovvalues"  or lcase(sAction)="verifypropertylabels" Or lcase(sAction)="verifypropertylabelsnotexist" or lcase(sAction)="seteditboxvalueandverifylength" or lcase(sAction)="seteditboxvalueandverifylengthequals" or lcase(sAction)="verifypropertynotexist" or lcase(sAction)="verifyonelovvalueallowtoselectfromlist" or Lcase(sAction)="verifyeditboxvalues" or Lcase(sAction)="verifymandatoryfields" Or Lcase(sAction)="verifylengthofid" Then

		DataTable.SetCurrentRow iItemCount
		If sItemType<>"" Then
			'Select Item type
			iItemNodeCount=Fn_UI_Object_Operations("RAC_Common_CreateItem","GetROProperty", objNewItem.JavaTree("jtree_ItemType"),"","items count","")
			For iCounter=0 To iItemNodeCount-1
				sCurrentItemName = objNewItem.JavaTree("jtree_ItemType").GetItem(iCounter)
				If Trim(sCurrentItemName)="Most Recently Used~" & Trim(sItemType) Then
					sItemType = "Most Recently Used~" & Trim(sItemType)
					bFlag=True
					Exit For
				ElseIf Trim(sCurrentItemName)="Complete List~" & Trim(sItemType) Then
					sItemType = "Complete List~" & Trim(sItemType)
					bFlag=True
					Exit For
				End If
			Next
			
			If bFlag = True Then
				If Fn_UI_JavaTree_Operations("RAC_Common_CreateItem","Select",objNewItem,"jtree_ItemType",sItemType,"","") = False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Item type [ " & Cstr(sItemType) & " ] from new item creation dialog","","","","","")
					Call Fn_ExitTest()
				End If
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to select Item type [ " & Cstr(sItemType) & " ] from new item creation dialog as specified item type does not exist in list","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
			
			'Click on next button
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_Next") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Next ] button from new item creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MIN_SYNC_ITERATIONS)
		End If
		sItemType=sTempItemType
	End If
Else
	sItemType=sTempItemType
End If

Select Case LCase(sAction)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic item with standard values
	Case "autobasiccreate"
		DataTable.SetCurrentRow iItemCount	
		'Assign item ID
		If Fn_UI_JavaButton_Operations("Fn_DatasetDetailsCreate", "Click", objNewItem,"jbtn_AssignID")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Assign ] button to assign id","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		'Assign revision
		If Fn_UI_JavaButton_Operations("Fn_DatasetDetailsCreate", "Click", objNewItem,"jbtn_AssignRevision")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Assign ] button to assign revision id","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Get auto generated Item ID
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","ID:")
		sItemID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "" )
		
		'Get revision ID
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Revision:")
		sItemRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "" )
				
		'Set Item Name
		sName = Fn_Setup_GenerateObjectInformation("getname",sItemType)
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Name:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
		'Set Item description		
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Description:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item description vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
		'Form the item node name in nav tree and store in Datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemNode","","")
		DataTable.Value("RACItemNode","Global") = sItemID & "-" & sName
		
		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemRevisionNode","","")
		DataTable.Value("RACItemRevisionNode","Global") = sItemID & "/" & sItemRevisionID & ";1-" & sName
		
		'Store Item ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemID","","")
		DataTable.Value("RACItemID","Global") = sItemID
		
		'Store Item Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemRevisionID","","")
		DataTable.Value("RACItemRevisionID","Global") = sItemRevisionID
		
		'Store Item Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemName","","")
		DataTable.Value("RACItemName","Global") = sName
				
		'Click on Finish button
		objNewItem.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_Cancel")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateItem",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Item of type [ " & Cstr(sItemType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created item of type [ " & Cstr(sItemType) & " ] with Item Id [ " & Cstr(Datatable.Value("RACItemID", "Global")) & " ] , Item Revision Id [ " & Cstr(Datatable.Value("RACItemRevisionID", "Global")) & " ] and Item name [ " & Cstr(Datatable.Value("RACItemName", "Global")) & " ]","","","","","")
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create basic item with custom values
	Case "basiccreate"
		DataTable.SetCurrentRow iItemCount	
		'Assign item ID
		If dictItemInfo("ID")="" or dictItemInfo("ID")="Assign" Then
			If Fn_UI_JavaButton_Operations("Fn_DatasetDetailsCreate", "Click", objNewItem,"jbtn_AssignID")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Assign ] button to assign id","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","ID:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", dictItemInfo("ID") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item id vlaue","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Assign revision
		If dictItemInfo("Revision")="" or dictItemInfo("Revision")="Assign" Then
			If Fn_UI_JavaButton_Operations("Fn_DatasetDetailsCreate", "Click", objNewItem,"jbtn_AssignRevision")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Assign ] button to assign revision id","","","","","")
				Call Fn_ExitTest()
			End If
		Else
			Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Revision:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", dictItemInfo("Revision") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item revision vlaue","","","","","")
				Call Fn_ExitTest()
			End If
		End If	
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Get auto generated Item ID
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","ID:")
		sItemID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "" )
		
		'Get revision ID
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Revision:")
		sItemRevisionID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "" )
				
		'Set Item Name
		If dictItemInfo("Name")="" or dictItemInfo("Name")="Assign" Then
			sName = Fn_Setup_GenerateObjectInformation("getname",sItemType)
		Else
			sName =dictItemInfo("Name")
		End If
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Name:")
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", sName )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item name vlaue","","","","","")
			Call Fn_ExitTest()
		End If									
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
					
		'Set Item description
		If dictItemInfo("Description")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Description:")
			If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", dictItemInfo("Description") )=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item description vlaue","","","","","")
				Call Fn_ExitTest()
			End If									
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
		
		'Set Item unit of measure
		If dictItemInfo("UnitOfMeasure")<>"" Then
			Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","Unit of Measure:")
			Call Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_LOVDropDown")
			Set objWScriptShell = CreateObject("WScript.Shell")
			wait GBL_MICRO_TIMEOUT
			objWScriptShell.SendKeys "{TAB}"
			wait GBL_MICRO_TIMEOUT
			objWScriptShell.SendKeys "{DOWN}"
			If objLOVTree.Exist(2)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item unit of measure vlaue","","","","","")
				Call Fn_ExitTest()
			End IF
			objLOVTree.Activate dictItemInfo("UnitOfMeasure")
			wait GBL_MIN_MICRO_TIMEOUT
			If Err.Number<>0 then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to set item unit of measure vlaue","","","","","")
				Call Fn_ExitTest()
			End If
		End If
		
		'Form the item node name in nav tree and store in Datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemNode","","")
		DataTable.Value("RACItemNode","Global") = sItemID & "-" & sName
		
		'Store nav tree revision node details in datatable
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemRevisionNode","","")
		DataTable.Value("RACItemRevisionNode","Global") = sItemID & "/" & sItemRevisionID & ";1-" & sName
		
		'Store Item ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemID","","")
		DataTable.Value("RACItemID","Global") = sItemID
		
		'Store Item Revision ID
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemRevisionID","","")
		DataTable.Value("RACItemRevisionID","Global") = sItemRevisionID
		
		'Store Item Name
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemName","","")
		DataTable.Value("RACItemName","Global") = sName
		
		'Store Item Unit of measure
		Call Fn_CommonUtil_DataTableOperations("AddColumn","RACItemUnitOfMeasure","","")
		DataTable.Value("RACItemUnitOfMeasure","Global") = dictItemInfo("UnitOfMeasure")
		
		'Click on Finish button
		objNewItem.JavaButton("jbtn_Finish").WaitProperty "enabled", 1, 20000
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_Finish")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Finish ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)

		'Click on Close button
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_Cancel")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new item of type [ " & Cstr(sItemType) & " ] as fail to click on [ Close ] button","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		Call Fn_LogUtil_CaptureFunctionExecutionTime("CaptureEndTime","RAC_Common_CreateItem",sAction,"","")
		If Err.Number<0 then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to create new Item of type [ " & Cstr(sItemType) & " ] due to error [ " & Cstr(Err.Description) & " ]","","","","","")
			Call Fn_ExitTest()
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_ACTION,"Successfully created item of type [ " & Cstr(sItemType) & " ] with Item Id [ " & Cstr(Datatable.Value("RACItemID", "Global")) & " ] , Item Revision Id [ " & Cstr(Datatable.Value("RACItemRevisionID", "Global")) & " ] and Item name [ " & Cstr(Datatable.Value("RACItemName", "Global")) & " ]","","","","","")
		End If
		DataTable.SetCurrentRow GBL_DATATABLEGLOBALSHEETCURRENTROW_NUMBER
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'case to click on any button on [ New Item ] creation dialog
	Case "clickbutton"
		If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	Case "verifylengthofid"
		'Assign item ID
		If Fn_UI_JavaButton_Operations("Fn_DatasetDetailsCreate", "Click", objNewItem,"jbtn_AssignID")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Assign ] button to assign id on [ New Item ] creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
		'Get auto generated Item ID
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label","ID:")
		sItemID = Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "" )
		If Len(Cstr(sItemID))=dictItemInfo("Length")Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified assign id length is [ " & Cstr(dictItemInfo("Length")) & " ]","","","","DONOTSYNC","")
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail :Verification fail as assign id length is not equal to [ " & Cstr(dictItemInfo("Length")) & " ]","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Create validate specific fields are mandetory
	Case "verifymandatoryfields"		
		If sItemType="Control Document" Then
			sFieldNames=Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL","ControlDocument_Mandatory_Fields",""))
		End If
		
		aFieldNames=Split(sFieldNames,"~")
		For iCounter=0 to Ubound(aFieldNames)	
			Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",aFieldNames(iCounter) & ":")
			If Fn_UI_Object_Operations("RAC_Common_CreateItem","Exist", objNewItem.JavaStaticText("jstx_Asterix"), GBL_DEFAULT_TIMEOUT,"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as field [ " & Cstr(aFieldNames(iCounter)) & " ] is not mandatory field for Item type [ " & Cstr(sItemType) & " ]","","","","","")
				Call Fn_ExitTest()
			Else
			    Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : Successfully verified fields [ " & Cstr(aFieldNames(iCounter)) & " ] are mandatory field for Item type [ " & Cstr(sItemType) & " ]","","","","DONOTSYNC","")
			End If
		Next
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ Close ] button of [ New Item Creation ] dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)	
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -		
	Case "verifyeditboxvalues"
		aProperty=Split(dictItemInfo("PropertyName"),"~")
		aValues=Split(dictItemInfo("PropertyValue"),"~")
				
		For iCounter=0 to UBound(aProperty)
			sTempPropertyName=aProperty(iCounter)
			aProperty(iCounter)=Fn_FSOUtil_XMLFileOperations("getvalue","RAC_ObjectPropertiesValues_APL",aProperty(iCounter),"")
			
			If Cstr(aProperty(iCounter))="False" Then
				aProperty(iCounter)=sTempPropertyName
			End If
			
			If aProperty(iCounter)<>"Relation" Then
				If Fn_UI_Object_Operations("RAC_Common_CreateItem","settoproperty",objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",aProperty(iCounter) & ":")=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on object creation dialog","","","","","")
					Call Fn_ExitTest()
				End IF
			Else
				If Fn_UI_Object_Operations("RAC_Common_CreateItem","settoproperty",objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",aProperty(iCounter))=False Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on object creation dialog","","","","","")
					Call Fn_ExitTest()
				End IF
			End IF
			
			If aValues(iCounter)="<<EMPTY>>" Then
				aValues(iCounter)=""
			End If
			
			If Fn_UI_Object_Operations("RAC_Common_CreateItem", "Exist", objNewItem.JavaEdit("jedt_ItemEdit"),"","","") Then
				If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "")=aValues(iCounter) Then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","DONOTSYNC","")
				Else
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not contain value [ " & Cstr(aValues(iCounter)) & " ]","","","","","")
					Call Fn_ExitTest()
				End If	
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(dictProperties("PropertyName")) & " ] property does not exist\available on object creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Verify only one LOV value is allow to select from list
	Case "verifyonelovvalueallowtoselectfromlist"
		objNewItem.Maximize
		dictItemInfo("LOVTreeLabel") = Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",dictItemInfo("LOVTreeLabel"),""))
	    Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",dictItemInfo("LOVTreeLabel") & ":")
		Call Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_LOVDropDown")
		
		'Verify existence of LOV tree
		If Fn_UI_Object_Operations("RAC_Common_CreateItem","Exist", objLOVTree,GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify values of [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree as [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree does not exist on new item creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		aLOVTreeNode=Split(Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",dictItemInfo("LOVTreeNode"),"")),"~")
	
		'aLOVTreeNode=Split(dictItemInfo("LOVTreeNode"),"~")		
		objLOVTree.SelectRange aLOVTreeNode(0),aLOVTreeNode(Ubound(aLOVTreeNode))
		Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		
		'Click on the dropdown button to close the lov tree
		Call Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_LOVDropDown")
		
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "GetText", objNewItem, "jedt_ItemEdit", "" )=aLOVTreeNode(0) Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : Successfully verified LOV list [ " & (dictItemInfo("LOVTreeLabel")) & " ] allows to select only one value from list","","","","DONOTSYNC","")	
		Else
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as LOV list [ " & (dictItemInfo("LOVTreeLabel")) & " ] allows to select multiple value from list","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "verifypropertylabels"   
		aProperty = Split(Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",dictItemInfo("PropertyLabel"),"")),"~")
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_CreateItem","settoexistcheck",objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on new item creation dialog","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property available on new item creation dialog","","","","DONOTSYNC","")
			End If
		Next	
		
        If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog while performing action [" & sAction & "]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "verifypropertylabelsnotexist"   
		aProperty = Split(Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",dictItemInfo("PropertyLabel"),"")),"~")
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_CreateItem","settoexistcheck",objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",aProperty(iCounter) & ":")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property does not exist\available on new item creation dialog","","","","DONOTSYNC","")				
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property exist\available on new item creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
		Next	
		
        If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog while performing action [" & sAction & "]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "verifybuttonenabled"
		 If sButton<>"" Then
        	If Fn_UI_Object_Operations("RAC_Common_CreateItem", "Enabled", objNewItem.JavaButton("jbtn_" & sButton),"", "", "") = False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as button [ " & Cstr(sButton) & " ] is not enabled on new item creation dialog","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as button [ " & Cstr(sButton) & " ] is enabled on new item creation dialog","","","","DONOTSYNC","")
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If		
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "verifybuttonnotenabled"			
        If sButton<>"" Then
        	If Fn_UI_Object_Operations("RAC_Common_CreateItem", "Enabled", objNewItem.JavaButton("jbtn_" & sButton),"", "", "") = True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as button [ " & Cstr(sButton) & " ] is enabled on new item creation dialog","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as button [ " & Cstr(sButton) & " ] is not enabled on new item creation dialog","","","","DONOTSYNC","")
			End If
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "seteditboxvalue", "seteditboxvalueandverifylength"  ,"seteditboxvalueandverifylengthequals"  
   		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label", dictItemInfo("PropertyLabel") & ":")		
		If Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "Set",  objNewItem, "jedt_ItemEdit", dictItemInfo("PropertyValue") )=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to set value of editbox [ " & Cstr(dictItemInfo("PropertyLabel")) & " ] as [" & dictItemInfo("PropertyValue") & "]","","","","","")
			Call Fn_ExitTest()
		End If
		
		If sAction = "seteditboxvalueandverifylength" Then
			If Cint(Len(Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "gettext",  objNewItem, "jedt_ItemEdit", "")))<> Cint(dictItemInfo("Length")) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as editbox [ " & Cstr(dictItemInfo("PropertyLabel")) & " ] has string value of length not equal to [" & dictItemInfo("Length") & "]","","","","DONOTSYNC","")	
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as editbox [ " & Cstr(dictItemInfo("PropertyLabel")) & " ] has string value of length equal to [" & dictItemInfo("Length") & "]","","","","","")
				Call Fn_ExitTest()
			End If
		ElseIf sAction = "seteditboxvalueandverifylengthequals" Then
			If Cint(Len(Fn_UI_JavaEdit_Operations("RAC_Common_CreateItem", "gettext",  objNewItem, "jedt_ItemEdit", "")))= Cint(dictItemInfo("Length")) Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Pass : verification pass as editbox [ " & Cstr(dictItemInfo("PropertyLabel")) & " ] has string value of length equal to [" & dictItemInfo("Length") & "]","","","","DONOTSYNC","")	
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as editbox [ " & Cstr(dictItemInfo("PropertyLabel")) & " ] has string value of length not equal to [" & dictItemInfo("Length") & "]","","","","","")
				Call Fn_ExitTest()
			End If
		End If

        If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog while performing action [" & sAction & "]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -	
   	Case "verifypropertynotexist"   
		aProperty = Split(Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",dictItemInfo("PropertyLabel"),"")),"~")
		For iCounter=0 to ubound(aProperty)
			If Fn_UI_Object_Operations("RAC_Common_CreateItem","settoexistcheck",objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",aProperty(iCounter) & ":")=True Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as [ " & Cstr(aProperty(iCounter)) & " ] property  exist\available on new item creation dialog","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified [ " & Cstr(aProperty(iCounter)) & " ] property does not available on new item creation dialog","","","","DONOTSYNC","")
			End If
		Next	
		
        If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog while performing action [" & sAction & "]","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'Verify LOV values		
	Case "verifylovvalues"
	    objNewItem.Maximize
		dictItemInfo("LOVTreeLabel") = Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",dictItemInfo("LOVTreeLabel"),""))
	    Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",dictItemInfo("LOVTreeLabel") & ":")
		objNewItem.JavaButton("jbtn_LOVDropDown").Highlight
		Call Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_LOVDropDown")
		
		'Verify existence of LOV tree
		If Fn_UI_Object_Operations("RAC_Common_CreateItem","Exist", objLOVTree,GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify values of [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree as [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree does not exist on new item creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		aLOVTreeNode=Split(dictItemInfo("LOVTreeNode"),"~")
		iItemCount=objLOVTree.GetROProperty("items count")
		
		For iCounter=0 to Ubound(aLOVTreeNode) 
		   sLOVTreeNode= Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",aLOVTreeNode(iCounter),""))
		   aName = Split(sLOVTreeNode,"~")
		   
		   If cint(iItemCount)=cint(Ubound(aName)+1)Then
			  For iNameCount = 0 To Ubound(aName) Step 1
				If trim(objLOVTree.GetItem(iNameCount))=trim(aName(iNameCount)) Then
				    Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified node [ " & Cstr(aName(iNameCount)) & " ] appears under [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree","","","","DONOTSYNC","")
				Else
				    Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify LOV tree node of [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree as node [ " & Cstr(aName(iNameCount)) & " ] does not available in LOV tree","","","","","")
					Call Fn_ExitTest()
				End If		
			  Next
		   End If	 
		Next
		
		'Click on the dropdown button to close the lov tree
		Call Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_LOVDropDown")
		
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -			
	'case to verify LOV tree node values
	Case "verifylovtreenode"
		Call Fn_UI_Object_Operations("RAC_Common_CreateItem","SetTOProperty", objNewItem.JavaStaticText("jstx_ItemLabel"),"","label",dictItemInfo("LOVTreeLabel") & ":")
		Call Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem, "jbtn_LOVDropDown")
		
		'Verify existence of LOV tree
		If Fn_UI_Object_Operations("RAC_Common_CreateItem","Exist", objNewItem.JavaTree("jtree_ItemName"),GBL_DEFAULT_TIMEOUT,"","")=False Then
			Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify values of [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree as [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree does not exist on new item creation dialog","","","","","")
			Call Fn_ExitTest()
		End If
		
		aLOVTreeNode=Split(dictItemInfo("LOVTreeNode"),"~")
		For iCounter=0 to Ubound(aLOVTreeNode)
			sLOVTreeNode = Cstr(Fn_FSOUtil_XMLFileOperations("getvalue","RAC_NewItemValues_APL",aLOVTreeNode(iCounter),""))
			'Expand the name tree till last parent hierarchy
			aName = Split(sLOVTreeNode,"~")
			For iNameCount = 0 To Ubound(aName) - 1 Step 1
				If iNameCount = 0 Then
					sNode = aName(iNameCount)
				Else
					sNode = sNode & "~" & aName(iNameCount)
				End If
				objNewItem.JavaTree("jtree_ItemName").Expand sNode
				Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
				
				If Err.Number<0 then
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to expand name node [ " & Cstr(sNode) & " ] due to error [ " & Cstr(Err.Description) & " ] while creating new Engineered Part","","","","","")
					Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify LOV tree node of [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree as node [ " & Cstr(sNode) & " ] does not available in LOV tree","","","","","")
					Call Fn_ExitTest()
				End If
			Next
			
			'If Fn_UI_JavaTree_Operations("RAC_MyTc_NavigationTreeOperations","Exist",objNewItem,"jtree_ItemName",sLOVTreeNode,"","")=False Then
			If Fn_RAC_GetJavaTreeNodeIndex(objNewItem.JavaTree("jtree_ItemName"),sLOVTreeNode,"","")=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : Fail to verify LOV tree node of [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree as node [ " & Cstr(sLOVTreeNode) & " ] does not available in LOV tree","","","","","")
				Call Fn_ExitTest()
			Else
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified node [ " & Cstr(sLOVTreeNode) & " ] appears under [ " & Cstr(dictItemInfo("LOVTreeLabel")) & " ] LOV tree","","","","DONOTSYNC","")
			End If		
		
		Next
		If sButton<>"" Then
			If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
				Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog","","","","","")
				Call Fn_ExitTest()
			End If
			Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
		End If	
	' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -                                            
     'Case to verify specific item type is not available under item type list
       Case "verifyitemtypenotexist"
            iItemNodeCount=Fn_UI_Object_Operations("RAC_Common_CreateItem","GetROProperty", objNewItem.JavaTree("jtree_ItemType"),"","items count","")
            For iCounter=0 To iItemNodeCount-1
                sCurrentItemName = objNewItem.JavaTree("jtree_ItemType").GetItem(iCounter)
                If Trim(sCurrentItemName)="Most Recently Used~" & Trim(sItemType) Then
                   sItemType = "Most Recently Used~" & Trim(sItemType)
                   bFlag=True
                   Exit For
                ElseIf Trim(sCurrentItemName)="Complete List~" & Trim(sItemType) Then
                   sItemType = "Complete List~" & Trim(sItemType)
                   bFlag=True
                   Exit For
                End If
            Next
                                
            If bFlag = True Then                                        
              Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Fail : verification fail as Item type [ " & Cstr(sItemType) & " ] appears under Item Type list for currently logged in user","","","","","")
              Call Fn_ExitTest()
            Else
              Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_PASS_VERIFICATION,"VP " &  GBL_VERIFICATION_COUNTER & " - Successfully verified Item type [ " & Cstr(sItemType) & " ] does not appears under Item Type list for currently logged in user","","","","DONOTSYNC","")
             End If
                                
            If sButton<>"" Then
                If Fn_UI_JavaButton_Operations("RAC_Common_CreateItem", "Click", objNewItem,"jbtn_" & sButton)=False Then
                   Call Fn_LogUtil_PrintAndUpdateScriptLog(GBL_LOG_FAIL_ACTION,"Fail to click on [ " & Cstr(sButton) & " ] button of new item creation dialog while performing action [" & sAction & "]","","","","","")
                   Call Fn_ExitTest()
                End If
                   Call Fn_RAC_ReadyStatusSync(GBL_MICRO_SYNC_ITERATIONS)
            End If
End Select

'Releasing object of New Item dialog
Set objNewItem =Nothing
Set objLOVTree=Nothing

Function Fn_ExitTest()
	'Releasing object of New Item dialog
	Set objNewItem =Nothing
	Set objLOVTree=Nothing
	ExitTest
End Function
