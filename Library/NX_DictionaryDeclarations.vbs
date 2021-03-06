Option Explicit

'Declaring variables
Dim dictNXItemInfo, dictNXSaveAsInfo, dictNXAssemblyInfo, dictNXWeightManagementInfo,dictExportPartInfo, dictImportPartInfo
Dim dictNXRoutingInfo

'to store information required to create Item in NX
Set dictNXItemInfo = CreateObject("Scripting.Dictionary")
'to store information required to save as Item in NX
Set dictNXSaveAsInfo = CreateObject("Scripting.Dictionary")
'to store information required to create new assembly in NX
Set dictNXAssemblyInfo = CreateObject("Scripting.Dictionary")
'to store information required to perform weight management in NX
Set dictNXWeightManagementInfo = CreateObject("Scripting.Dictionary")
'to store information required to perform export part operations in NX
Set dictExportPartInfo = CreateObject("Scripting.Dictionary")
'to store information required to perform import part operations in NX
Set dictImportPartInfo = CreateObject("Scripting.Dictionary")
'to store information required to perform routing operation from Teamcenter to NX
Set dictNXRoutingInfo = CreateObject("Scripting.Dictionary")