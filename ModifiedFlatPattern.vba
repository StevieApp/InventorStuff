Sub NewMacro()
    Dim oPartDoc As PartDocument
    Set oPartDoc = ThisApplication.ActiveDocument
    Dim oDef As ControlDefinition
    Set oDef = ThisApplication.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
    oDef.Execute
    Dim oCompDef As SheetMetalComponentDefinition
    Set oCompDef = oPartDoc.ComponentDefinition
    ''If Not oCompDef.HasFlatPattern Then
    oCompDef.Unfold
    Dim oFlatPtn As FlatPattern
    Set oFlatPtn = oCompDef.FlatPattern
    ''MsgBox (oFlatPtn.FlatBendResults.Count)
    ''Set oDef = ThisApplication.CommandManager.ControlDefinitions.Item("PartSwitchRepresentationCmd")
    ''oDef.Execute
    Dim oUOM As UnitsOfMeasure
    Set oUOM = ThisApplication.ActiveDocument.UnitsOfMeasure
    Dim oDTProps As PropertySet
    PropertySet = oPartDoc.PropertySets.Item("Design Tracking Properties")
    MsgBox (oFlatPtn.MassProperties.Area & " in " & oUOM.GetStringFromType(oUOM.LengthUnits))
End Sub
