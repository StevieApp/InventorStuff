Sub PartsFromBOM()
    Dim oAsmDoc As AssemblyDocument
    Set oAsmDoc = ThisApplication.ActiveDocument
    Dim oBOM As BOM
    Set oBOM = oAsmDoc.ComponentDefinition.BOM
    oBOM.PartsOnlyViewEnabled = True
    Dim oPartsOnlyBOMView As BOMView
    Set oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")
    Dim i As Long
    For i = 1 To oPartsOnlyBOMView.BOMRows.Count
        Dim oRow As BOMRow
        Set oRow = oPartsOnlyBOMView.BOMRows.Item(i)
        Dim oCompDef As ComponentDefinition
        Set oCompDef = oRow.ComponentDefinitions.Item(1)
        Dim oRefDoc As Document
        Set oRefDoc = oRow.ComponentDefinitions.Item(1).Document
        If TypeName(oCompDef) = "SheetMetalComponentDefinition" Then
            Dim oCompDeff As SheetMetalComponentDefinition
            Set oCompDeff = oCompDef
            Dim oFlatPtn As FlatPattern
            Set oFlatPtn = oCompDeff.FlatPattern
            oCompDeff.Unfold
            Dim invCustomPropertySet As PropertySet
            Set invCustomPropertySet = oRefDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim dblValue, db2Value As String
            dblValue = Round(oFlatPtn.MassProperties.Area * 100, 2) & " sq. mm"
            db2Value = Round(oRow.ItemQuantity * oFlatPtn.MassProperties.Area * 100, 2) & " sq. mm"
            On Error Resume Next
            Dim invVolumeProperty, invVolumeProperty0 As Property
            Set invVolumeProperty = invCustomPropertySet.Item("Flat Sheet Area")
            Set invVolumeProperty0 = invCustomPropertySet.Item("Total Flat Sheet Area")
            If Err.Number <> 0 Then
                Call invCustomPropertySet.Add(dblValue, "Flat Sheet Area")
                Call invCustomPropertySet.Add(db2Value, "Total Flat Sheet Area")
            Else
                invVolumeProperty.Value = dblValue
                invVolumeProperty0.Value = db2Value
            End If
        Else
            Dim invCustomPropertySett As PropertySet
            Set invCustomPropertySett = oRefDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim dblValuee, db2Valuee As String
            dblValuee = "Not Applicable"
            db2Valuee = "Not Applicable"
            On Error Resume Next
            Dim invVolumePropertyy, invVolumePropertyy0 As Property
            Set invVolumePropertyy = invCustomPropertySett.Item("Flat Sheet Area")
            Set invVolumePropertyy0 = invCustomPropertySett.Item("Total Flat Sheet Area")
            If Err.Number <> 0 Then
                Call invCustomPropertySett.Add(dblValuee, "Flat Sheet Area")
                Call invCustomPropertySett.Add(db2Valuee, "Total Flat Sheet Area")
            Else
                invVolumePropertyy.Value = dblValuee
                invVolumePropertyy0.Value = db2Valuee
            End If
        End If
    Next
End Sub
