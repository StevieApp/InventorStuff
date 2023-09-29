Sub DivideTrial()
    Dim oAsmDoc As AssemblyDocument
    Set oAsmDoc = ThisApplication.ActiveDocument
    Dim oRefDocs As DocumentsEnumerator
    Set oRefDocs = oAsmDoc.AllReferencedDocuments
    Dim oRefDoc As Document
    Dim i
    i = 0
    Dim Dict As Scripting.Dictionary
    Set Dict = New Scripting.Dictionary
    For Each oRefDoc In oRefDocs
        If oRefDoc.DocumentType = kPartDocumentObject Then
            ''oMsg = oMsg & vbLf & Left("Part: " & oRefDoc.DisplayName & " Count: " & i, Len("Part: " & oRefDoc.DisplayName & " Count: " & i))
            Dim Milo As PartDocument
            Set Milo = oRefDoc
            If Dict.Exists(Milo.DisplayName) Then
            Else
                i = i + 1
                If TypeName(Milo.ComponentDefinition) = "SheetMetalComponentDefinition" Then
                    Dim oCompDef As SheetMetalComponentDefinition
                    Set oCompDef = Milo.ComponentDefinition
                    oCompDef.Unfold
                    Dim oFlatPtn As FlatPattern
                    Set oFlatPtn = oCompDef.FlatPattern
                    Dim mystring
                    mystring = "Sheetmetal: " & Milo.DisplayName & " " & oFlatPtn.MassProperties.Area * 100 & " Area in sq. mm " & " Count: " & i
                    oMsg = oMsg & vbLf & Left(mystring, Len(mystring))
                    Dict.Add Key:=Milo.DisplayName, Item:=oFlatPtn.MassProperties.Area * 100
                    Dim invCustomPropertySet As PropertySet
                    Set invCustomPropertySet = oRefDoc.PropertySets.Item("Inventor User Defined Properties")
                    Dim dblValue As String
                    dblValue = Round(oFlatPtn.MassProperties.Area * 100, 2) & " sq. mm"
                    On Error Resume Next
                    Dim invVolumeProperty As Property
                    Set invVolumeProperty = invCustomPropertySet.Item("Sheet Area")
                    If Err.Number <> 0 Then
                        Call invCustomPropertySet.Add(dblValue, "Sheet Area")
                    Else
                        invVolumeProperty.Value = dblValue
                    End If
                    ''Set invProperty = invCustomPropertySet.Add(oFlatPtn.MassProperties.Area, "MyFlatArea")
                Else
                    Dim mestring
                    mestring = "Part: " & Milo.DisplayName & " N/A sq. mm" & " Count: " & i
                    oMsg = oMsg & vbLf & Left(mestring, Len(mestring))
                    Dict.Add Key:=Milo.DisplayName, Item:="N/A"
                    Dim invCustomPropertySett As PropertySet
                    Set invCustomPropertySett = oRefDoc.PropertySets.Item("Inventor User Defined Properties")
                    Dim dblValuee As String
                    dblValuee = "N/A"
                    On Error Resume Next
                    Dim invVolumePropertyy As Property
                    Set invVolumePropertyy = invCustomPropertySett.Item("Sheet Area")
                    If Err.Number <> 0 Then
                        Call invCustomPropertySett.Add(dblValuee, "Sheet Area")
                    Else
                        invVolumePropertyy.Value = dblValuee
                    End If
                End If
            End If
        ElseIf oRefDoc.DocumentType = kAssemblyDocumentObject Then
            ''oMsg = oMsg & vbLf & Left("Assembly " & oRefDoc.DisplayName, Len("Assembly " & oRefDoc.DisplayName) - 2)
            Call Further(oRefDoc, oMsg, i, Dict)
        End If
    Next
    MsgBox (oMsg)
    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile("C:\Users\Proteq_Automation\Videos\mojo\sagome.txt", True)
    a.WriteLine (oMsg)
    a.Close
    Dim oBOM As BOM
    Set oBOM = oAsmDoc.ComponentDefinition.BOM
    oBOM.PartsOnlyViewEnabled = True
    Dim oPartsOnlyBOMView As BOMView
    Set oPartsOnlyBOMView = oBOM.BOMViews.Item("Parts Only")
    ''Dim i As Long
    ''For i = 1 To oBOMView.BOMRows.Count

        ''Dim oRow As BOMRow
        ''Set oRow = oBOMView.BOMRows.Item(i)
    ''Next
    ''oBOM.BOMViews.
    ''ThisDoc.Launch ("C:\Users\Proteq_Automation\Videos\mojo\sagome.txt")
End Sub

Sub Further(Sembly, oMsg, i, Dict)
    Dim oAsmDoccy As AssemblyDocument
    Set oAsmDoccy = Sembly
    Dim oRefDocz As DocumentsEnumerator
    Set oRefDocz = oAsmDoccy.AllReferencedDocuments
    Dim oRefDocc As Document
    For Each oRefDocc In oRefDocz
        If oRefDocc.DocumentType = kPartDocumentObject Then
            Dim Miloo As PartDocument
            Set Miloo = oRefDocc
            If Dict.Exists(Miloo.DisplayName) Then
            Else
                i = i + 1
                If TypeName(Miloo.ComponentDefinition) = "SheetMetalComponentDefinition" Then
                    Dim oCompDef As SheetMetalComponentDefinition
                    Set oCompDef = Miloo.ComponentDefinition
                    oCompDef.Unfold
                    Dim oFlatPtn As FlatPattern
                    Set oFlatPtn = oCompDef.FlatPattern
                    Dim mystringg
                    mystringg = "Sheetmetal: " & Miloo.DisplayName & " " & oFlatPtn.MassProperties.Area * 100 & " Area in sq. mm" & " Count: " & i
                    oMsg = oMsg & vbLf & Left(mystringg, Len(mystringg))
                    Dict.Add Key:=Miloo.DisplayName, Item:=oFlatPtn.MassProperties.Area * 100
                    Dim invCustomPropertySet As PropertySet
                    Set invCustomPropertySet = oRefDocc.PropertySets.Item("Inventor User Defined Properties")
                    Dim dblValue As String
                    dblValue = Round(oFlatPtn.MassProperties.Area * 100, 2) & " sq. mm"
                    On Error Resume Next
                    Dim invVolumeProperty As Property
                    Set invVolumeProperty = invCustomPropertySet.Item("Sheet Area")
                    If Err.Number <> 0 Then
                        Call invCustomPropertySet.Add(dblValue, "Sheet Area")
                    Else
                        invVolumeProperty.Value = dblValue
                    End If
                Else
                    Dim mestringg
                    mestringg = "Part: " & Miloo.DisplayName & " N/A sq. mm" & " Count: " & i
                    oMsg = oMsg & vbLf & Left(mestringg, Len(mestringg))
                    Dict.Add Key:=Miloo.DisplayName, Item:="N/A"
                    Dim invCustomPropertySett As PropertySet
                    Set invCustomPropertySett = oRefDocc.PropertySets.Item("Inventor User Defined Properties")
                    Dim dblValuee As String
                    dblValuee = "N/A"
                    On Error Resume Next
                    Dim invVolumePropertyy As Property
                    Set invVolumePropertyy = invCustomPropertySett.Item("Sheet Area")
                    If Err.Number <> 0 Then
                        Call invCustomPropertySett.Add(dblValuee, "Sheet Area")
                    Else
                        invVolumePropertyy.Value = dblValuee
                    End If
                End If
            End If
        ElseIf oRefDocc.DocumentType = kAssemblyDocumentObject Then
            ''oMsg = oMsg & vbLf & Left("Assembly " & oRefDoc.DisplayName, Len("Assembly " & oRefDoc.DisplayName) - 2)
            Call Further(oRefDocc, oMsg, i, Dict)
        End If
    Next
End Sub
