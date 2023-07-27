Sub Main
 Dim oDoc As PartDocument
    Set oDoc = ThisApplication.ActiveDocument
    Dim oDef As PartComponentDefinition
    Set oDef = oDoc.ComponentDefinition
    Dim LengthTotal As Double
    LengthTotal = 0
    'add length of all sweeps
    Dim oSweep As SweepFeature
    MsgBox (oDef.Features.SweepFeatures.Count)
    For Each oSweep In oDef.Features.SweepFeatures
        LengthTotal = LengthTotal + GetLengthOfSweep(oSweep.Path)
        MsgBox (LengthTotal)
    Next oSweep
    'get(or create) parameter
    Dim oPara As Parameter
    Set oPara = GetPara(oDoc.ComponentDefinition.Parameters, "SweepLength")
    'assign value and comment
    oPara.Value = LengthTotal
    Dim oComment As String
    oComment = "Determined by iLogic-rule: 'SweepLength'"
    If Not oPara.Comment = oComment Then oPara.Comment = oComment
    oDoc.Update
End Sub

'get length of a sweep
Private Function GetLengthOfSweep(ByVal oPath As Path) As Double
    Dim oCurve As Object
    Dim oCurveEval As CurveEvaluator
    Dim MinParam As Double
    Dim MaxParam As Double
    Dim length As Double
    Dim TotalLength As Double
    TotalLength = 0
    i As Integer
    For i = 1 To oPath.Count
        oCurve = oPath.Item(i).Curve
        oCurveEval = oCurve.Evaluator
        Call oCurveEval.GetParamExtents(MinParam, MaxParam)
        Call oCurveEval.GetLengthAtParam(MinParam, MaxParam, length)
        TotalLength = TotalLength + length
    Next i
    GetLengthOfSweep = TotalLength
End Function

Private Function GetPara(ByRef oParas As Parameters, ByVal paraName As String) As Parameter
    Dim oPara As Parameter
    For Each oPara In oParas
        If (oPara.Name = paraName) Then Set GetPara = oPara
    Next oPara
    Set GetPara = oParas.UserParameters.AddByValue(paraName, 0, 11266)
End Function
