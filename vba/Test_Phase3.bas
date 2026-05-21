Attribute VB_Name = "Test_Phase3"
' Test_Phase3 - branch coverage skeleton and call graph tests.
' Run via TestRunner.Run_All_Tests.

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_BranchEdges_If"
    TestRunner.Run_One "Test_BranchEdges_Evaluate_NoOther"
    TestRunner.Run_One "Test_CallGraph_Edges"
    TestRunner.Run_One "Test_ICASE1_TotalBranches"
End Sub

Public Sub Test_BranchEdges_If()
    Dim src As String
    src = "000100 PROCEDURE DIVISION." & vbCrLf & _
          "000200     IF X = 1" & vbCrLf & _
          "000300         MOVE 'A' TO Y" & vbCrLf & _
          "000400     ELSE" & vbCrLf & _
          "000500         MOVE 'B' TO Y" & vbCrLf & _
          "000600     END-IF."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim roots As Collection
    Set roots = CobolParser.Get_CobolNodes(norm.Item("Lines"))
    Dim edges As Collection
    Set edges = CobolParser.Get_BranchEdges(roots)
    TestRunner.Assert_Equal CLng(2), CLng(edges.Count), "IF yields 2 edges"
    TestRunner.Assert_Equal "if", CStr(edges.Item(1).Item("type")), "edge1 type=if"
End Sub

Public Sub Test_BranchEdges_Evaluate_NoOther()
    Dim src As String
    src = "000100 PROCEDURE DIVISION." & vbCrLf & _
          "000200     EVALUATE M" & vbCrLf & _
          "000300         WHEN 1" & vbCrLf & _
          "000400             MOVE 'A' TO Y" & vbCrLf & _
          "000500         WHEN 2" & vbCrLf & _
          "000600             MOVE 'B' TO Y" & vbCrLf & _
          "000700     END-EVALUATE."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim roots As Collection
    Set roots = CobolParser.Get_CobolNodes(norm.Item("Lines"))
    Dim edges As Collection
    Set edges = CobolParser.Get_BranchEdges(roots)
    ' 2 WHEN edges + 1 nomatch edge (no OTHER given) = 3
    TestRunner.Assert_Equal CLng(3), CLng(edges.Count), "EVALUATE 2 WHEN -> 3 edges"
End Sub

Public Sub Test_CallGraph_Edges()
    Dim src As String
    src = "000100 PROCEDURE DIVISION." & vbCrLf & _
          "000200 MAIN-PARA." & vbCrLf & _
          "000300     PERFORM SUB-PARA." & vbCrLf & _
          "000400     CALL 'EXTPRG'." & vbCrLf & _
          "000500 SUB-PARA." & vbCrLf & _
          "000600     MOVE 1 TO X."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim st As OrderedDict
    Set st = CobolParser.Get_ProgramStructure(norm.Item("Lines"))
    Dim cg As OrderedDict
    Set cg = CobolParser.Get_CallRelationships(norm.Item("Lines"), st)
    Dim edges As Collection
    Set edges = cg.Item("edges")
    TestRunner.Assert_True CLng(edges.Count) >= 2, "at least PERFORM + CALL edges"
    Dim hasPerform As Boolean, hasCall As Boolean, e As OrderedDict
    For Each e In edges
        If e.Item("kind") = "perform" Then hasPerform = True
        If e.Item("kind") = "call" Then hasCall = True
    Next e
    TestRunner.Assert_True hasPerform, "perform edge present"
    TestRunner.Assert_True hasCall, "call edge present"
End Sub

Public Sub Test_ICASE1_TotalBranches()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(cblPath)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & cblPath
        Exit Sub
    End If
    Dim src As String
    src = CobolEncoding.ReadAllText(cblPath, "auto")
    Dim r As OrderedDict
    Set r = CobolParser.Analyze_Phase3(src)
    Dim cov As OrderedDict
    Set cov = r.Item("coverage")
    ' Golden ICASE1: totalBranches=21
    TestRunner.Assert_Equal CLng(21), CLng(cov.Item("totalBranches")), "ICASE1 totalBranches=21"
End Sub
