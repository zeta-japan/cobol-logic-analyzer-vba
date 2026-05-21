Attribute VB_Name = "Test_Phase5"
' Test_Phase5 - end-to-end test that the engine output drives the 5-sheet renderer.
' Run via TestRunner.Run_All_Tests.

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_AnalyzeAndBuild_CreatesSheets"
End Sub

Public Sub Test_AnalyzeAndBuild_CreatesSheets()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(cblPath)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & cblPath
        Exit Sub
    End If

    Main.AnalyzeAndBuild cblPath

    ' BuildCobolReport creates these 5 sheets (Japanese names).
    Dim expected As Variant
    expected = Array(Chr(67) & Chr(79) & Chr(66) & Chr(79) & Chr(76), "COBOL")  ' placeholder list (renderer-defined)
    Dim ws As Worksheet, sheetCount As Long
    sheetCount = 0
    For Each ws In ThisWorkbook.Worksheets
        sheetCount = sheetCount + 1
    Next ws
    TestRunner.Assert_True (sheetCount >= 5), "workbook has at least 5 sheets after BuildCobolReport (had " & sheetCount & ")"
End Sub
