Attribute VB_Name = "modTPLogger"
Option Explicit

Public Sub CreateNewTestPitSheet()
    Dim wsTemplate As Worksheet
    Dim wsIndex As Worksheet
    Dim newName As String
    Dim wsNew As Worksheet
    Dim nextRow As Long

    Set wsTemplate = ThisWorkbook.Worksheets("TP_Template")
    Set wsIndex = ThisWorkbook.Worksheets("Index")

    newName = InputBox("Enter new sheet name, e.g. TP01", "Create New Test Pit")
    If Trim$(newName) = "" Then Exit Sub

    If WorksheetExists(newName) Then
        MsgBox "Sheet already exists: " & newName, vbExclamation
        Exit Sub
    End If

    wsTemplate.Visible = xlSheetVisible
    wsTemplate.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    Set wsNew = ActiveSheet
    wsNew.Name = newName
    wsTemplate.Visible = xlSheetVeryHidden

    wsNew.Range("B4").Value = "TP"
    wsNew.Range("B5").Value = newName
    wsNew.Range("B6").Value = ThisWorkbook.Worksheets("ProjectInfo").Range("B3").Value
    wsNew.Range("B15").Value = ThisWorkbook.Worksheets("ProjectInfo").Range("B10").Value

    nextRow = wsIndex.Cells(wsIndex.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 4 Then nextRow = 4

    wsIndex.Cells(nextRow, 1).Value = nextRow - 3
    wsIndex.Cells(nextRow, 2).Value = "TP"
    wsIndex.Cells(nextRow, 3).Value = newName
    wsIndex.Cells(nextRow, 4).Value = newName
    wsIndex.Cells(nextRow, 14).Value = "Draft"

    MsgBox "Created new sheet: " & newName, vbInformation
End Sub

Public Sub RefreshIndexFromSheets()
    Dim wsIndex As Worksheet
    Dim ws As Worksheet
    Dim nextRow As Long

    Set wsIndex = ThisWorkbook.Worksheets("Index")
    wsIndex.Range("A4:N1000").ClearContents
    nextRow = 4

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "ProjectInfo" And ws.Name <> "Index" And ws.Name <> "TP_Template" _
           And ws.Name <> "LookupTables" And ws.Name <> "Samples" And ws.Name <> "Summary" _
           And ws.Name <> "CrossSectionData" And ws.Name <> "Export_All" Then

            wsIndex.Cells(nextRow, 1).Value = nextRow - 3
            wsIndex.Cells(nextRow, 2).Value = Nz(ws.Range("B4").Value)
            wsIndex.Cells(nextRow, 3).Value = ws.Name
            wsIndex.Cells(nextRow, 4).Value = Nz(ws.Range("B5").Value)
            wsIndex.Cells(nextRow, 5).Value = ws.Range("B7").Value
            wsIndex.Cells(nextRow, 6).Value = ws.Range("B8").Value
            wsIndex.Cells(nextRow, 7).Value = Nz(ws.Range("B9").Value)
            wsIndex.Cells(nextRow, 8).Value = Nz(ws.Range("B10").Value)
            wsIndex.Cells(nextRow, 9).Value = ws.Range("B12").Value
            wsIndex.Cells(nextRow, 10).Value = ws.Range("B13").Value
            wsIndex.Cells(nextRow, 11).Value = ws.Range("B14").Value
            wsIndex.Cells(nextRow, 12).Value = ws.Range("B15").Value
            wsIndex.Cells(nextRow, 13).Value = BuildTerminationText(ws)
            wsIndex.Cells(nextRow, 14).Value = "Updated"
            nextRow = nextRow + 1
        End If
    Next ws

    MsgBox "Index refreshed.", vbInformation
End Sub

Public Sub RebuildSamples()
    Dim wsSamples As Worksheet
    Dim ws As Worksheet
    Dim r As Long, outRow As Long

    Set wsSamples = ThisWorkbook.Worksheets("Samples")
    wsSamples.Range("A4:H1000").ClearContents
    outRow = 4

    For Each ws In ThisWorkbook.Worksheets
        If IsPointSheet(ws) Then
            For r = 25 To 34
                If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "sample" Then
                    wsSamples.Cells(outRow, 1).Value = Nz(ws.Cells(r, 3).Value)
                    wsSamples.Cells(outRow, 2).Value = Nz(ws.Range("B5").Value)
                    wsSamples.Cells(outRow, 3).Value = ws.Name
                    wsSamples.Cells(outRow, 4).Value = Nz(ws.Cells(r, 4).Value)
                    wsSamples.Cells(outRow, 5).Value = ws.Cells(r, 5).Value
                    wsSamples.Cells(outRow, 6).Value = ws.Cells(r, 6).Value
                    If IsNumeric(ws.Cells(r, 5).Value) And IsNumeric(ws.Cells(r, 6).Value) Then
                        wsSamples.Cells(outRow, 7).Value = (CDbl(ws.Cells(r, 5).Value) + CDbl(ws.Cells(r, 6).Value)) / 2
                    End If
                    wsSamples.Cells(outRow, 8).Value = Nz(ws.Cells(r, 2).Value)
                    outRow = outRow + 1
                End If
            Next r
        End If
    Next ws

    MsgBox "Samples rebuilt.", vbInformation
End Sub

Public Function BuildTerminationText(ByVal ws As Worksheet) As String
    Dim r As Long
    For r = 25 To 34
        If LCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = "termination" Then
            BuildTerminationText = Nz(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
    BuildTerminationText = ""
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
    IsPointSheet = Not (ws.Name = "ProjectInfo" Or ws.Name = "Index" Or ws.Name = "TP_Template" _
        Or ws.Name = "LookupTables" Or ws.Name = "Samples" Or ws.Name = "Summary" _
        Or ws.Name = "CrossSectionData" Or ws.Name = "Export_All")
End Function
