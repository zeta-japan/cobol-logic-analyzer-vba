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

    Dim beforeCount As Long, afterCount As Long, ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        beforeCount = beforeCount + 1
    Next ws

    Main.AnalyzeAndBuild cblPath

    For Each ws In ThisWorkbook.Worksheets
        afterCount = afterCount + 1
    Next ws
    TestRunner.Assert_True (afterCount >= 5), _
        "workbook has at least 5 sheets after BuildCobolReport (before=" & beforeCount & ", after=" & afterCount & ")"
End Sub
