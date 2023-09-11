
Sub DivideTrial()
    Dim oAsmDoc As AssemblyDocument
    Set oAsmDoc = ThisApplication.ActiveDocument
    Dim oRefDocs As DocumentsEnumerator
    Set oRefDocs = oAsmDoc.AllReferencedDocuments
    Dim oRefDoc As Document
    Dim i
    i = 0
    For Each oRefDoc In oRefDocs
        If oRefDoc.DocumentType = kPartDocumentObject Then
            i = i + 1
            oMsg = oMsg & vbLf & Left("Part: " & oRefDoc.DisplayName & " Count: " & i, Len("Part: " & oRefDoc.DisplayName & " Count: " & i))
        ElseIf oRefDoc.DocumentType = kAssemblyDocumentObject Then
            ''oMsg = oMsg & vbLf & Left("Assembly " & oRefDoc.DisplayName, Len("Assembly " & oRefDoc.DisplayName) - 2)
            Call Further(oRefDoc, oMsg, i)
        End If
    Next
    MsgBox (oMsg)
End Sub

Sub Further(Sembly, oMsg, i)
    Dim oAsmDoccy As AssemblyDocument
    Set oAsmDoccy = Sembly
    Dim oRefDocz As DocumentsEnumerator
    Set oRefDocz = oAsmDoccy.AllReferencedDocuments
    Dim oRefDocc As Document
    For Each oRefDocc In oRefDocz
        If oRefDocc.DocumentType = kPartDocumentObject Then
            i = i + 1
            oMsg = oMsg & vbLf & Left("Part: " & oRefDocc.DisplayName & " Count: " & i, Len("Part: " & oRefDocc.DisplayName & " Count: " & i))
        ElseIf oRefDocc.DocumentType = kAssemblyDocumentObject Then
            ''oMsg = oMsg & vbLf & Left("Assembly " & oRefDoc.DisplayName, Len("Assembly " & oRefDoc.DisplayName) - 2)
            Call Further(oRefDocc, oMsg, i)
        End If
    Next
End Sub
