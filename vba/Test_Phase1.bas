Attribute VB_Name = "Test_Phase1"
' Test_Phase1 - smoke tests for the Phase 1 deliverable.
' Run via TestRunner.Run_All_Tests.

Option Explicit

Public Sub Run_All()
    Test_OrderedDict_KeysOrder
    Test_JsonWriter_Primitives
    Test_JsonWriter_OrderedDict
    Test_JsonWriter_Collection
    Test_Convert_CollapseSpaces
    Test_Convert_StripTrailingPeriod
    Test_NormalizedLines_PrefixDetected
    Test_ProgramName_FromIcase1
End Sub

Private Sub Test_OrderedDict_KeysOrder()
    Dim od As OrderedDict
    Set od = New OrderedDict
    od.Add "first", 1
    od.Add "second", "two"
    od.Add "third", True
    TestRunner.Assert_Equal 3, od.Count, "OrderedDict.Count"
    TestRunner.Assert_Equal "first", od.Keys.Item(1), "Keys[1]=first"
    TestRunner.Assert_Equal "second", od.Keys.Item(2), "Keys[2]=second"
    TestRunner.Assert_Equal "third", od.Keys.Item(3), "Keys[3]=third"
    TestRunner.Assert_Equal 1, od.Item("first"), "Item(first)"
    TestRunner.Assert_Equal "two", od.Item("second"), "Item(second)"
End Sub

Private Sub Test_JsonWriter_Primitives()
    TestRunner.Assert_Equal "null", JsonWriter.WriteJson(Null), "WriteJson Null"
    TestRunner.Assert_Equal "true", JsonWriter.WriteJson(True), "WriteJson True"
    TestRunner.Assert_Equal "false", JsonWriter.WriteJson(False), "WriteJson False"
    TestRunner.Assert_Equal "42", JsonWriter.WriteJson(CLng(42)), "WriteJson 42"
    TestRunner.Assert_Equal """hello""", JsonWriter.WriteJson("hello"), "WriteJson hello"
    TestRunner.Assert_Equal """a\""b""", JsonWriter.WriteJson("a" & Chr$(34) & "b"), "WriteJson quote escape"
End Sub

Private Sub Test_JsonWriter_OrderedDict()
    Dim od As OrderedDict
    Set od = New OrderedDict
    od.Add "name", "ICASE1"
    od.Add "lines", CLng(100)
    TestRunner.Assert_Equal "{""name"":""ICASE1"",""lines"":100}", JsonWriter.WriteJson(od), "WriteJson OrderedDict order"
End Sub

Private Sub Test_JsonWriter_Collection()
    Dim c As Collection
    Set c = New Collection
    c.Add CLng(1)
    c.Add "two"
    c.Add True
    TestRunner.Assert_Equal "[1,""two"",true]", JsonWriter.WriteJson(c), "WriteJson Collection"
End Sub

Private Sub Test_Convert_CollapseSpaces()
    TestRunner.Assert_Equal "A B C", CobolParser.Convert_CollapseSpaces("  A   B  C  "), "Convert_CollapseSpaces"
End Sub

Private Sub Test_Convert_StripTrailingPeriod()
    TestRunner.Assert_Equal "IF X = 1", CobolParser.Convert_StripTrailingPeriod("IF X = 1."), "StripTrailingPeriod"
End Sub

Private Sub Test_NormalizedLines_PrefixDetected()
    Dim src As String
    src = "000100 IDENTIFICATION DIVISION." & vbCrLf & _
          "000200 PROGRAM-ID. SMOKE." & vbCrLf & _
          "000300 PROCEDURE DIVISION." & vbCrLf & _
          "000400 MAIN-SECTION SECTION."
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    TestRunner.Assert_Equal "prefixed", norm.Item("PrefixStyle"), "PrefixStyle auto"
    TestRunner.Assert_True CBool(norm.Item("PrefixDetected")), "PrefixDetected=True"
End Sub

Private Sub Test_ProgramName_FromIcase1()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    Dim src As String
    src = CobolEncoding.ReadAllText(cblPath, "auto")
    Dim result As OrderedDict
    Set result = CobolParser.Analyze_Phase1(src)
    Dim summary As OrderedDict
    Set summary = result.Item("summary")
    TestRunner.Assert_Equal "ICASE1", summary.Item("programName"), "ICASE1 programName"
    TestRunner.Assert_True (CLng(summary.Item("lines")) > 0), "ICASE1 lines > 0"
End Sub
