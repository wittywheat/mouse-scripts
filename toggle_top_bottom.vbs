
Option Explicit

' Add any type libraries to be used.
Scripting.AddTypeLibrary("MGCPCB.ExpeditionPCBApplication")

' Get the application object.
Dim pcbAppObj
Set pcbAppObj = Application

' Get the active document
Dim pcbDocObj
Set pcbDocObj = pcbAppObj.ActiveDocument

' License the document
ValidateServer(pcbDocObj)

' Find out how many copper layers are there in this PCB
Dim layerObjColl
Set layerObjColl = pcbDocObj.LayerStack(False)

Dim layer_count
layer_count = layerObjColl.Count

' Toggle between top and bottom layers
Dim pcbGuiObj
Set pcbGuiObj = pcbAppObj.Gui

Dim active_layer
active_layer = pcbGuiObj.ActiveRouteLayer

If active_layer <> 1 Then
	pcbGuiObj.ActiveRouteLayer = 1
Else
	pcbGuiObj.ActiveRouteLayer = layer_count
End If


' Server validation function
Function ValidateServer(docObj)
	Dim keyInt
	Dim licenseTokenInt
	Dim licenseServerObj

	' Ask Expedition’s document for the key
	keyInt = docObj.Validate(0)

	' Get license server
	Set licenseServerObj = CreateObject("MGCPCBAutomationLicensing.Application")

	' Ask the license server for the license token
	licenseTokenInt = licenseServerObj.GetToken(keyInt)

	' Release license server
	Set licenseServerObj = nothing

	' Turn off error messages (validate may fail if the token is incorrect)
	On Error Resume Next
	Err.Clear

	' Ask the document to validate the license token
	docObj.Validate(licenseTokenInt)
	If Err Then
		ValidateServer = 0
	Else
		ValidateServer = 1
	End If

End Function