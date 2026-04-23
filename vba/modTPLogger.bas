Attribute VB_Name = "modTPLogger"
Option Explicit

Private Const SH_PROJECT As String = "ProjectInfo"
Private Const SH_INDEX As String = "Index"
Private Const SH_TEMPLATE As String = "TP_Template"
Private Const SH_SAMPLES As String = "Samples"
Private Const SH_SUMMARY As String = "Summary"
Private Const SH_XS As String = "CrossSectionData"
Private Const SH_EXPORT As String = "Export_All"
Private Const SH_LOOKUPS As String = "LookupTables"

Public Sub CreateNewTestPitSheet()
    Dim wsTemplate As Worksheet
    Dim wsIndex As Worksheet
    Dim wsNew As Worksheet
    Dim newName As String
    Dim nextRow As Long

    Set wsTemplate = ThisWorkbook.Worksheets(SH_TEMPLATE)
    Set wsIndex = ThisWorkbook.Worksheets(SH_INDEX)

    newName = Trim$(InputBox("Enter new sheet name, e.g. TP01", "Create New Test Pit"))
    If newName = "" Then Exit Sub

    If WorksheetExists(newName) Then
        MsgBox "Sheet already exists: " & newName, vbExclamation
        Exit Sub
    End If

    wsTemplate.Visible = xlSheetVisible
    wsTemplate.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set wsNew = ActiveSheet
    wsNew.Name = newName
    wsTemplate.Visible = xlSheetVeryHidden

    PopulateProjectDefaults wsNew, newName

    nextRow = NextIndexRow(wsIndex)
    wsIndex.Cells(nextRow, 1).Value = nextRow - 3
    wsIndex.Cells(nextRow, 2).Value = Nz(wsNew.Range("B4").Value)
    wsIndex.Cells(nextRow, 3).Value = wsNew.Name
    wsIndex.Cells(nextRow, 4).Value = Nz(wsNew.Range("B5").Value)
    wsIndex.Cells(nextRow, 14).Value = "Draft"

    MsgBox "Created new sheet: " & newName, vbInformation
End Sub

Public Sub RefreshCurrentSheet()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If Not IsPointSheet(ws) Then
        MsgBox "Active sheet is not a point sheet.", vbExclamation
        Exit Sub
    End If
    AutoCalculateLayerFields ws
    BuildLayerPreviews ws
    BuildHolePreview ws
    HighlightValidation ws
    RefreshIndexFromSheets
    RebuildSamples
    MsgBox ws.Name & " refreshed.", vbInformation
End Sub

Public Sub RefreshAll()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsPointSheet(ws) Then
            AutoCalculateLayerFields ws
            BuildLayerPreviews ws
            BuildHolePreview ws
            HighlightValidation ws
        End If
    Next ws
    RefreshIndexFromSheets
    RebuildSamples
    BuildCrossSectionData
    BuildCombinedExportPreview
    MsgBox "All sheets refreshed.", vbInformation
End Sub

Public Sub RefreshIndexFromSheets()
    Dim wsIndex As Worksheet, ws As Worksheet, nextRow As Long
    Set wsIndex = ThisWorkbook.Worksheets(SH_INDEX)
    wsIndex.Range("A4:N1000").ClearContents
    nextRow = 4
    For Each ws In ThisWorkbook.Worksheets
        If IsPointSheet(ws) Then
            wsIndex.Cells(nextRow, 1).Value = nextRow - 3
            wsIndex.Cells(nextRow, 2).Value = Nz(ws.Range("B4").Value)
            wsIndex.Cells(nextRow, 3).Value = ws.Name
            wsIndex.Cells(nextRow, 4).Value = Nz(ws.Range("B5").Value)
            wsIndex.Cells(nextRow, 5).Value = ws.Range("B7").Value
            wsIndex.Cells(nextRow, 6).Value = ws.Range("B8").Value
            wsIndex.Cells(nextRow, 7).Value = Nz(ws.Range("B9").Value)
            wsIndex.Cells(nextRow, 8).Value = Nz(ws.Range("B10").Value)
            wsIndex.Cells(nextRow, 9).Value = ws.Range("B13").Value
            wsIndex.Cells(nextRow, 10).Value = ws.Range("B14").Value
            wsIndex.Cells(nextRow, 11).Value = ws.Range("B15").Value
            wsIndex.Cells(nextRow, 12).Value = ws.Range("B16").Value
            wsIndex.Cells(nextRow, 13).Value = BuildTerminationText(ws)
            wsIndex.Cells(nextRow, 14).Value = "Updated"
            nextRow = nextRow + 1
        End If
    Next ws
End Sub

Public Sub RebuildSamples()
    Dim wsSamples As Worksheet, ws As Worksheet
    Dim r As Long, outRow As Long
    Set wsSamples = ThisWorkbook.Worksheets(SH_SAMPLES)
    wsSamples.Range("A4:H1000").ClearContents
    outRow = 4
    For Each ws In ThisWorkbook.Worksheets
        If IsPointSheet(ws) Then
            For r = 25 To 36
                If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "sample" Then
                    wsSamples.Cells(outRow, 1).Value = Nz(ws.Cells(r, 3).Value)
                    wsSamples.Cells(outRow, 2).Value = Nz(ws.Range("B5").Value)
                    wsSamples.Cells(outRow, 3).Value = ws.Name
                    wsSamples.Cells(outRow, 4).Value = Nz(ws.Cells(r, 4).Value)
                    wsSamples.Cells(outRow, 5).Value = ws.Cells(r, 5).Value
                    wsSamples.Cells(outRow, 6).Value = ws.Cells(r, 6).Value
                    If IsNumeric(ws.Cells(r, 5).Value) And IsNumeric(ws.Cells(r, 6).Value) Then
                        wsSamples.Cells(outRow, 7).Value = Round((CDbl(ws.Cells(r, 5).Value) + CDbl(ws.Cells(r, 6).Value)) / 2, 2)
                    End If
                    wsSamples.Cells(outRow, 8).Value = Nz(ws.Cells(r, 2).Value)
                    outRow = outRow + 1
                End If
            Next r
        End If
    Next ws
End Sub

Public Sub BuildCrossSectionData()
    Dim wsOut As Worksheet, ws As Worksheet
    Dim r As Long, outRow As Long
    Dim rl As Double, fromD As Double, toD As Double
    Set wsOut = ThisWorkbook.Worksheets(SH_XS)
    wsOut.Range("A4:J5000").ClearContents
    outRow = 4
    For Each ws In ThisWorkbook.Worksheets
        If IsPointSheet(ws) Then
            rl = SafeDbl(ws.Range("B15").Value)
            For r = 5 To 20
                If SafeDbl(ws.Cells(r, 5).Value) > 0 Or SafeDbl(ws.Cells(r, 6).Value) > 0 Then
                    fromD = SafeDbl(ws.Cells(r, 5).Value)
                    toD = SafeDbl(ws.Cells(r, 6).Value)
                    wsOut.Cells(outRow, 1).Value = Nz(ws.Range("B5").Value)
                    wsOut.Cells(outRow, 2).Value = ws.Name
                    wsOut.Cells(outRow, 3).Value = ws.Range("B13").Value
                    wsOut.Cells(outRow, 4).Value = ws.Range("B14").Value
                    wsOut.Cells(outRow, 5).Value = rl
                    wsOut.Cells(outRow, 6).Value = fromD
                    wsOut.Cells(outRow, 7).Value = toD
                    wsOut.Cells(outRow, 8).Value = rl - fromD
                    wsOut.Cells(outRow, 9).Value = rl - toD
                    wsOut.Cells(outRow, 10).Value = Nz(ws.Cells(r, 17).Value)
                    outRow = outRow + 1
                End If
            Next r
        End If
    Next ws
End Sub

Public Sub BuildCombinedExportPreview()
    Dim wsOut As Worksheet, ws As Worksheet
    Dim outRow As Long, r As Long
    Set wsOut = ThisWorkbook.Worksheets(SH_EXPORT)
    wsOut.Range("A4:D5000").ClearContents
    outRow = 4
    For Each ws In ThisWorkbook.Worksheets
        If IsPointSheet(ws) Then
            wsOut.Cells(outRow, 1).Value = outRow - 3
            wsOut.Cells(outRow, 2).Value = ws.Name
            wsOut.Cells(outRow, 3).Value = Nz(ws.Range("B5").Value)
            wsOut.Cells(outRow, 4).Value = "JOB NAME: " & Nz(ProjectValue(2))
            outRow = outRow + 1
            wsOut.Cells(outRow, 4).Value = "/"
            outRow = outRow + 1
            For r = 5 To 20
                If Nz(ws.Cells(r, 18).Value) <> "" Then
                    wsOut.Cells(outRow, 4).Value = Format$(SafeDbl(ws.Cells(r, 5).Value), "0.00") & vbTab & Nz(ws.Cells(r, 18).Value)
                    outRow = outRow + 1
                End If
            Next r
            wsOut.Cells(outRow, 4).Value = "/"
            outRow = outRow + 1
            wsOut.Cells(outRow, 4).Value = "NOTES:"
            outRow = outRow + 1
            For r = 25 To 36
                If Nz(ws.Cells(r, 1).Value) <> "" And Nz(ws.Cells(r, 2).Value) <> "" Then
                    wsOut.Cells(outRow, 4).Value = Nz(ws.Cells(r, 1).Value) & ": " & Nz(ws.Cells(r, 2).Value)
                    outRow = outRow + 1
                End If
            Next r
            wsOut.Cells(outRow, 4).Value = "//"
            outRow = outRow + 2
        End If
    Next ws
    wsOut.Cells(outRow, 4).Value = "///"
End Sub

Public Sub BuildHolePreview(ByVal ws As Worksheet)
    Dim outRow As Long, r As Long
    ws.Range("J25:J36").ClearContents
    outRow = 25
    ws.Cells(outRow, 10).Value = "JOB NAME: " & Nz(ProjectValue(2))
    outRow = outRow + 1
    ws.Cells(outRow, 10).Value = "/"
    outRow = outRow + 1
    For r = 5 To 20
        If Nz(ws.Cells(r, 18).Value) <> "" Then
            If outRow <= 36 Then
                ws.Cells(outRow, 10).Value = Format$(SafeDbl(ws.Cells(r, 5).Value), "0.00") & vbTab & Nz(ws.Cells(r, 18).Value)
                outRow = outRow + 1
            End If
        End If
    Next r
End Sub

Public Sub AutoCalculateLayerFields(ByVal ws As Worksheet)
    Dim r As Long, fromD As Double, toD As Double
    Dim sampleFrom As Double, sampleTo As Double
    For r = 5 To 20
        fromD = SafeDbl(ws.Cells(r, 5).Value)
        toD = SafeDbl(ws.Cells(r, 6).Value)
        If toD > fromD Then
            ws.Cells(r, 7).Value = Round(toD - fromD, 2)
        Else
            ws.Cells(r, 7).ClearContents
        End If
    Next r
    For r = 25 To 36
        sampleFrom = SafeDbl(ws.Cells(r, 5).Value)
        sampleTo = SafeDbl(ws.Cells(r, 6).Value)
        If sampleTo > sampleFrom Then
            ws.Cells(r, 7).Value = Round((sampleFrom + sampleTo) / 2, 2)
        ElseIf sampleFrom > 0 Then
            ws.Cells(r, 7).Value = sampleFrom
        End If
    Next r
End Sub

Public Sub BuildLayerPreviews(ByVal ws As Worksheet)
    Dim r As Long
    For r = 5 To 20
        ws.Cells(r, 17).Value = BuildLayerDescription(ws, r)
        ws.Cells(r, 18).Value = BuildDotPlotLine(ws, r)
    Next r
End Sub

Public Function BuildLayerDescription(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    Dim moisture As String, colour1 As String, colour2 As String
    Dim consDen As String, structureDesc As String, soilType As String
    Dim origin As String, materialType As String, layerNote As String
    Dim desc As String, mainPart As String

    moisture = Nz(ws.Cells(rowNum, 8).Value)
    colour1 = Nz(ws.Cells(rowNum, 9).Value)
    colour2 = Nz(ws.Cells(rowNum, 10).Value)
    consDen = Nz(ws.Cells(rowNum, 11).Value)
    structureDesc = Nz(ws.Cells(rowNum, 12).Value)
    soilType = Nz(ws.Cells(rowNum, 13).Value)
    origin = Nz(ws.Cells(rowNum, 14).Value)
    materialType = Nz(ws.Cells(rowNum, 15).Value)
    layerNote = Nz(ws.Cells(rowNum, 16).Value)

    If Nz(ws.Cells(rowNum, 5).Value) = "" And Nz(ws.Cells(rowNum, 6).Value) = "" Then
        BuildLayerDescription = ""
        Exit Function
    End If

    mainPart = JoinComma(moisture, JoinColour(colour1, colour2), consDen, structureDesc, soilType)
    If origin <> "" Then
        If mainPart <> "" Then
            desc = mainPart & "; " & origin
        Else
            desc = origin
        End If
    Else
        desc = mainPart
    End If

    If materialType <> "" Then
        If desc <> "" Then
            desc = desc & ". " & materialType
        Else
            desc = materialType
        End If
    End If

    If layerNote <> "" Then
        If desc <> "" Then
            desc = desc & ". " & layerNote
        Else
            desc = layerNote
        End If
    End If

    If desc <> "" And Right$(desc, 1) <> "." Then desc = desc & "."
    BuildLayerDescription = desc
End Function

Public Function BuildDotPlotLine(ByVal ws As Worksheet, ByVal rowNum As Long) As String
    BuildDotPlotLine = BuildLayerDescription(ws, rowNum)
End Function

Public Sub HighlightValidation(ByVal ws As Worksheet)
    Dim r As Long, soilType As String, consDen As String
    ws.Range("K5:K20").Interior.Color = RGB(255, 242, 204)
    For r = 5 To 20
        soilType = LCase$(Nz(ws.Cells(r, 13).Value))
        consDen = LCase$(Nz(ws.Cells(r, 11).Value))
        If (SafeDbl(ws.Cells(r, 5).Value) > 0 Or SafeDbl(ws.Cells(r, 6).Value) > 0) Then
            If SafeDbl(ws.Cells(r, 6).Value) <= SafeDbl(ws.Cells(r, 5).Value) Then ws.Cells(r, 6).Interior.Color = RGB(255, 199, 206)
            If Nz(ws.Cells(r, 13).Value) = "" Then ws.Cells(r, 13).Interior.Color = RGB(255, 199, 206)
        End If
        If IsNonCohesiveSoil(soilType) And IsCohesiveDescriptor(consDen) Then
            ws.Cells(r, 11).Interior.Color = RGB(255, 235, 156)
        ElseIf IsCohesiveSoil(soilType) And IsNonCohesiveDescriptor(consDen) Then
            ws.Cells(r, 11).Interior.Color = RGB(255, 235, 156)
        End If
    Next r
End Sub

Public Function BuildTerminationText(ByVal ws As Worksheet) As String
    Dim r As Long
    For r = 25 To 36
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "termination" Then
            BuildTerminationText = Nz(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    BuildTerminationText = ""
End Function

Public Function NextIndexRow(ByVal wsIndex As Worksheet) As Long
    NextIndexRow = wsIndex.Cells(wsIndex.Rows.Count, 1).End(xlUp).Row + 1
    If NextIndexRow < 4 Then NextIndexRow = 4
End Function

Public Sub PopulateProjectDefaults(ByVal ws As Worksheet, ByVal newName As String)
    ws.Range("B4").Value = "TP"
    ws.Range("B5").Value = newName
    ws.Range("B6").Value = ProjectValue(1)
    ws.Range("B9").Value = ProjectValue(9)
    ws.Range("B10").Value = ProjectValue(10)
    ws.Range("B17").Value = ProjectValue(8)
End Sub

Public Function ProjectValue(ByVal projectRow As Long) As Variant
    ProjectValue = ThisWorkbook.Worksheets(SH_PROJECT).Cells(projectRow + 2, 2).Value
End Function

Public Function JoinComma(ParamArray parts() As Variant) As String
    Dim i As Long, outText As String, partText As String
    For i = LBound(parts) To UBound(parts)
        partText = Nz(parts(i))
        If partText <> "" Then
            If outText <> "" Then outText = outText & ", "
            outText = outText & partText
        End If
    Next i
    JoinComma = outText
End Function

Public Function JoinColour(ByVal colour1 As String, ByVal colour2 As String) As String
    If colour1 <> "" And colour2 <> "" Then
        JoinColour = colour1 & ", " & colour2
    Else
        JoinColour = JoinComma(colour1, colour2)
    End If
End Function

Public Function IsCohesiveDescriptor(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "very soft", "soft", "firm", "stiff", "very stiff": IsCohesiveDescriptor = True
    End Select
End Function

Public Function IsNonCohesiveDescriptor(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "very loose", "loose", "medium dense", "dense", "very dense": IsNonCohesiveDescriptor = True
    End Select
End Function

Public Function IsNonCohesiveSoil(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    IsNonCohesiveSoil = (InStr(s, "sand") > 0 Or InStr(s, "gravel") > 0 Or InStr(s, "cobble") > 0 Or InStr(s, "boulder") > 0) And InStr(s, "clay") = 0 And InStr(s, "silt") = 0
    If s = "silty sand" Or s = "clayey sand" Then IsNonCohesiveSoil = True
End Function

Public Function IsCohesiveSoil(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    IsCohesiveSoil = (InStr(s, "clay") > 0 Or InStr(s, "silt") > 0)
End Function

Public Function Nz(ByVal v As Variant) As String
    If IsError(v) Then
        Nz = ""
    ElseIf IsNull(v) Then
        Nz = ""
    ElseIf Trim$(CStr(v)) = "" Then
        Nz = ""
    Else
        Nz = Trim$(CStr(v))
    End If
End Function

Public Function SafeDbl(ByVal v As Variant) As Double
    If IsNumeric(v) Then SafeDbl = CDbl(v) Else SafeDbl = 0#
End Function

Public Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    WorksheetExists = False
    For Each ws In ThisWorkbook.Worksheets
        If LCase$(ws.Name) = LCase$(sheetName) Then
            WorksheetExists = True
            Exit Function
        End If
    Next ws
End Function

Public Function IsPointSheet(ByVal ws As Worksheet) As Boolean
    IsPointSheet = Not (ws.Name = SH_PROJECT Or ws.Name = SH_INDEX Or ws.Name = SH_TEMPLATE Or ws.Name = SH_LOOKUPS Or ws.Name = SH_SAMPLES Or ws.Name = SH_SUMMARY Or ws.Name = SH_XS Or ws.Name = SH_EXPORT)
End Function

Public Sub RunSmokeTest()
    On Error GoTo FailHandler
    CreateNewTestPitSheet
    RefreshAll
    MsgBox "Smoke test completed.", vbInformation
    Exit Sub
FailHandler:
    MsgBox "Smoke test failed: " & Err.Description, vbCritical
End Sub
