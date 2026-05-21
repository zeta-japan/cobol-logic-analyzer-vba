Attribute VB_Name = "Test_Phase4"
' Test_Phase4 - path enumeration and test case generation.
' Run via TestRunner.Run_All_Tests.

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_ExpandSequence_LinearActions"
    TestRunner.Run_One "Test_ExpandSequence_IfElse"
    TestRunner.Run_One "Test_ICASE1_Full_Counts"
    TestRunner.Run_One "Test_ICASE1_Full_Coverage100"
End Sub

Public Sub Test_ExpandSequence_LinearActions()
    Dim src As String
    src = "000100 PROCEDURE DIVISION." & vbCrLf & _
          "000200     MOVE 1 TO X." & vbCrLf & _
          "000300     MOVE 2 TO Y."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim roots As Collection
    Set roots = CobolParser.Get_CobolNodes(norm.Item("Lines"))
    Dim e As Collection
    Set e = New Collection
    Dim e2 As Collection: Set e2 = New Collection
    Dim e3 As Collection: Set e3 = New Collection
    Dim e4 As Collection: Set e4 = New Collection
    Dim states As Collection
    Set states = CobolParser.Expand_NodeSequence(roots, e, e2, e3, e4)
    TestRunner.Assert_Equal CLng(1), CLng(states.Count), "1 linear path"
    Dim st As OrderedDict
    Set st = states.Item(1)
    TestRunner.Assert_Equal CLng(2), CLng(st.Item("actions").Count), "2 actions in path"
End Sub

Public Sub Test_ExpandSequence_IfElse()
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
    Dim e1 As Collection: Set e1 = New Collection
    Dim e2 As Collection: Set e2 = New Collection
    Dim e3 As Collection: Set e3 = New Collection
    Dim e4 As Collection: Set e4 = New Collection
    Dim states As Collection
    Set states = CobolParser.Expand_NodeSequence(roots, e1, e2, e3, e4)
    TestRunner.Assert_Equal CLng(2), CLng(states.Count), "IF/ELSE -> 2 paths"
End Sub

Public Sub Test_ICASE1_Full_Counts()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(cblPath)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & cblPath
        Exit Sub
    End If
    Dim src As String
    src = CobolEncoding.ReadAllText(cblPath, "auto")
    Dim r As OrderedDict
    Set r = CobolParser.Analyze_Full(src)
    Dim s As OrderedDict
    Set s = r.Item("summary")
    ' Golden ICASE1: pathCount=96
    TestRunner.Assert_Equal CLng(96), CLng(s.Item("pathCount")), "ICASE1 pathCount=96"
    TestRunner.Assert_True Not CBool(s.Item("pathTruncated")), "ICASE1 not truncated"
    TestRunner.Assert_Equal CLng(96), CLng(r.Item("testCases").Count), "ICASE1 testCases=96"
End Sub

Public Sub Test_ICASE1_Full_Coverage100()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(cblPath)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & cblPath
        Exit Sub
    End If
    Dim src As String
    src = CobolEncoding.ReadAllText(cblPath, "auto")
    Dim r As OrderedDict
    Set r = CobolParser.Analyze_Full(src)
    Dim cov As OrderedDict
    Set cov = r.Item("coverage")
    TestRunner.Assert_Equal CLng(21), CLng(cov.Item("totalBranches")), "totalBranches=21"
    TestRunner.Assert_Equal CLng(21), CLng(cov.Item("coveredBranches")), "coveredBranches=21"
End Sub
