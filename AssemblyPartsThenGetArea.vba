Sub GetDerivedParts()
    Dim oApp As Application

    Dim oPD As AssemblyDocument

    Dim oDerPart As DerivedPartComponent
    
    Dim oRefDoc As Document
    
    Set oApp = ThisApplication

    Set oPD = oApp.ActiveDocument

    For Each oRefDoc In oPD.ReferencedDocuments
        If oRefDoc.DocumentType = kPartDocumentObject Then
            Dim partName As String
            For Each oDerPart In oRefDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents
                partName = oDerPart.Name
            Next
            Dim oPartDoc As partDocument
            Set oPartDoc = oRefDoc
            Dim oDef As ControlDefinition
            Set oDef = ThisApplication.CommandManager.ControlDefinitions.Item("PartConvertToSheetMetalCmd")
            oDef.Execute
            Dim oCompDef As SheetMetalComponentDefinition
            Set oCompDef = oPartDoc.ComponentDefinition
            If Not oCompDef.HasFlatPattern Then oCompDef.Unfold
            Dim oFlatPtn As FlatPattern
            Set oFlatPtn = oCompDef.FlatPattern
            MsgBox (partName & " " & oFlatPtn.MassProperties.Area & "sq. mm with " & oFlatPtn.FlatBendResults.Count & " bends")
        End If
    Next
End Sub
