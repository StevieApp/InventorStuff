Sub NewMacro()
    Dim oPartDoc As PartDocument
    Set oPartDoc = ThisApplication.ActiveDocument
    Dim oDef As ControlDefinition
    Set oDef = ThisApplication.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
    oDef.Execute
    Dim oCompDef As SheetMetalComponentDefinition
    Set oCompDef = oPartDoc.ComponentDefinition
    If Not oCompDef.HasFlatPattern Then oCompDef.Unfold
    Dim oFlatPtn As FlatPattern
    Set oFlatPtn = oCompDef.FlatPattern
    MsgBox (oFlatPtn.FlatBendResults.Count)
End Sub
'' refold
Set oDef = ThisApplication.CommandManager.ControlDefinitions.Item("PartSwitchRepresentationCmd")
oDef.Execute
