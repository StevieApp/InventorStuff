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
            ''MsgBox (TypeName(oPartDoc.ComponentDefinition))
            Set oCompDef = oPartDoc.ComponentDefinition
            If Not oCompDef.HasFlatPattern Then oCompDef.Unfold
            Dim oFlatPtn As FlatPattern
            Set oFlatPtn = oCompDef.FlatPattern
            MsgBox (partName & " " & oFlatPtn.MassProperties.Area & "sq. mm with " & oFlatPtn.FlatBendResults.Count & " bends")
        ElseIf oRefDoc.DocumentType = kAssemblyDocumentObject Then
            Call EnterAssembly(oRefDoc)
        End If
    Next
End Sub

Private Function EnterAssembly(Assembly As Document)
    Dim oApp As Application
    Dim oPD As AssemblyDocument
    Dim oDerPart As DerivedPartComponent
    Dim oRefDoc As Document
    Set oPD = Assembly
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
            If TypeName(oPartDoc.ComponentDefinition) = "SheetMetalComponentDefinition" Then
                Set oCompDef = oPartDoc.ComponentDefinition
                If Not oCompDef.HasFlatPattern Then oCompDef.Unfold
                Dim oFlatPtn As FlatPattern
                Set oFlatPtn = oCompDef.FlatPattern
                MsgBox (partName & " " & oFlatPtn.MassProperties.Area & "sq. mm with " & oFlatPtn.FlatBendResults.Count & " bends")
            ElseIf TypeName(oPartDoc.ComponentDefinition) = "PartComponentDefinition" Then
                MsgBox (partName & " is still a part")
            Else
            MsgBox (partName & " is neither a part nor sheetmetal")
            End If
        End If
    Next
End Function
