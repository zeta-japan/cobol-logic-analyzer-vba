Attribute VB_Name = "Test_Phase2"
' Test_Phase2 - AST construction and programStructure tests.
' Run via TestRunner.Run_All_Tests.

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_InvertCondition"
    TestRunner.Run_One "Test_GetCobolAction_Perform"
    TestRunner.Run_One "Test_GetCobolAction_NonAction"
    TestRunner.Run_One "Test_ParagraphName_Excluded"
    TestRunner.Run_One "Test_ProgramStructure_Sections"
    TestRunner.Run_One "Test_GetCobolNodes_SimpleIf"
    TestRunner.Run_One "Test_NodeCount_Recursion"
    TestRunner.Run_One "Test_ICASE1_StructureCounts"
End Sub

Public Sub Test_InvertCondition()
    TestRunner.Assert_Equal "NOT (X = 1)", CobolParser.Convert_InvertCondition("X = 1"), "Invert X=1"
    TestRunner.Assert_Equal "NOT (OTHER)", CobolParser.Convert_InvertCondition("OTHER"), "Invert OTHER"
End Sub

Public Sub Test_GetCobolAction_Perform()
    Dim n As OrderedDict
    Set n = CobolParser.Get_CobolAction("PERFORM ICASE1-INIT", 44)
    TestRunner.Assert_True Not n Is Nothing, "PERFORM returns non-Nothing"
    TestRunner.Assert_Equal "action", n.Item("type"), "PERFORM type=action"
    TestRunner.Assert_Equal "action-44", n.Item("id"), "PERFORM id"
    TestRunner.Assert_Equal "PERFORM ICASE1-INIT", n.Item("label"), "PERFORM label"
    TestRunner.Assert_Equal CLng(44), CLng(n.Item("startLine")), "PERFORM startLine"
End Sub

Public Sub Test_GetCobolAction_NonAction()
    Dim n As OrderedDict
    Set n = CobolParser.Get_CobolAction("IDENTIFICATION DIVISION", 1)
    TestRunner.Assert_True n Is Nothing, "non-action returns Nothing"
End Sub

Public Sub Test_ParagraphName_Excluded()
    TestRunner.Assert_True Not CobolParser.Test_ParagraphName("ELSE"), "ELSE excluded"
    TestRunner.Assert_True Not CobolParser.Test_ParagraphName("END-IF"), "END-IF excluded"
    TestRunner.Assert_True CobolParser.Test_ParagraphName("ICASE1-INIT"), "ICASE1-INIT allowed"
End Sub

Public Sub Test_ProgramStructure_Sections()
    Dim src As String
    src = "000100 IDENTIFICATION DIVISION." & vbCrLf & _
          "000200 PROGRAM-ID. SAMPLE." & vbCrLf & _
          "000300 PROCEDURE DIVISION." & vbCrLf & _
          "000400 MAIN-SECTION SECTION." & vbCrLf & _
          "000500 INIT-PARA." & vbCrLf & _
          "000600     PERFORM A-PARA THRU B-PARA." & vbCrLf & _
          "000700 END-PROGRAM."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim st As OrderedDict
    Set st = CobolParser.Get_ProgramStructure(norm.Item("Lines"))
    TestRunner.Assert_Equal CLng(1), CLng(st.Item("sections").Count), "sections=1"
    TestRunner.Assert_True CLng(st.Item("paragraphs").Count) >= 1, "paragraphs >= 1"
    TestRunner.Assert_Equal CLng(1), CLng(st.Item("performThruRanges").Count), "performThru=1"
End Sub

Public Sub Test_GetCobolNodes_SimpleIf()
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
    TestRunner.Assert_Equal CLng(1), CLng(roots.Count), "1 root node (IF)"
    Dim ifn As OrderedDict
    Set ifn = roots.Item(1)
    TestRunner.Assert_Equal "if", CStr(ifn.Item("type")), "type=if"
    TestRunner.Assert_Equal "X = 1", CStr(ifn.Item("condition")), "condition"
    TestRunner.Assert_Equal CLng(1), CLng(ifn.Item("thenChildren").Count), "1 thenChild"
    TestRunner.Assert_Equal CLng(1), CLng(ifn.Item("elseChildren").Count), "1 elseChild"
End Sub

Public Sub Test_NodeCount_Recursion()
    Dim src As String
    src = "000100 PROCEDURE DIVISION." & vbCrLf & _
          "000200     IF X = 1" & vbCrLf & _
          "000300         IF Y = 2" & vbCrLf & _
          "000400             MOVE 1 TO Z" & vbCrLf & _
          "000500         END-IF" & vbCrLf & _
          "000600     END-IF."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim roots As Collection
    Set roots = CobolParser.Get_CobolNodes(norm.Item("Lines"))
    TestRunner.Assert_Equal CLng(2), CLng(CobolParser.Get_NodeCount(roots, Array("if"))), "2 IF total"
    TestRunner.Assert_Equal CLng(1), CLng(CobolParser.Get_NodeCount(roots, Array("action"))), "1 action total"
End Sub

Public Sub Test_ICASE1_StructureCounts()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(cblPath)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & cblPath
        Exit Sub
    End If
    Dim src As String
    src = CobolEncoding.ReadAllText(cblPath, "auto")
    Dim r As OrderedDict
    Set r = CobolParser.Analyze_Phase2(src)
    Dim s As OrderedDict
    Set s = r.Item("summary")
    ' Expected values from the golden JSON for ICASE1.cbl.
    TestRunner.Assert_Equal CLng(14), CLng(s.Item("branchCount")), "ICASE1 branchCount=14"
    TestRunner.Assert_Equal CLng(36), CLng(s.Item("actionCount")), "ICASE1 actionCount=36"
    TestRunner.Assert_Equal CLng(8), CLng(s.Item("sectionCount")), "ICASE1 sectionCount=8"
    TestRunner.Assert_Equal CLng(20), CLng(s.Item("paragraphCount")), "ICASE1 paragraphCount=20"
    TestRunner.Assert_Equal CLng(1), CLng(s.Item("performThruCount")), "ICASE1 performThruCount=1"
End Sub
