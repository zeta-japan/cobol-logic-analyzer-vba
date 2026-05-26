Attribute VB_Name = "TestRunner"
' TestRunner - lightweight in-workbook test framework.
' Provides Assert_True / Assert_Equal helpers and renders results to a
' TestResults sheet plus the Immediate window. Run_All_Tests is the entry.

Option Explicit

Private mPass As Long
Private mFail As Long
Private mLog As Collection

Public Sub Test_Begin()
    mPass = 0
    mFail = 0
    Set mLog = New Collection
End Sub

Public Sub Assert_True(ByVal cond As Boolean, ByVal msg As String)
    If cond Then
        mPass = mPass + 1
        Log_ "[PASS] " & msg
    Else
        mFail = mFail + 1
        Log_ "[FAIL] " & msg
    End If
End Sub

Public Sub Assert_Equal(ByVal expected As Variant, ByVal actual As Variant, ByVal msg As String)
    Dim equal As Boolean
    On Error Resume Next
    equal = (expected = actual)
    On Error GoTo 0
    If equal Then
        mPass = mPass + 1
        Log_ "[PASS] " & msg
    Else
        mFail = mFail + 1
        Log_ "[FAIL] " & msg & "  expected=" & SafeToString_(expected) & "  actual=" & SafeToString_(actual)
    End If
End Sub

Public Sub Test_End()
    Log_ "---- " & mPass & " passed, " & mFail & " failed"
    Render_Log
End Sub

' Fast engine test suite (Phase 1-4). Does NOT render sheets, so it stays
' quick and does not mutate the workbook. Run this for routine verification.
Public Sub Run_All_Tests()
    Test_Begin
    Log_ "Workbook path: " & ThisWorkbook.path
    Test_Phase1.Run_All
    Test_Phase2.Run_All
    Test_Phase3.Run_All
    Test_Phase4.Run_All
    Test_Phase6.Run_All
    Test_Phase7.Run_All
    Test_Phase8.Run_All
    Test_End
End Sub

' Heavy end-to-end test: runs the full pipeline AND renders the 5 sheets via
' BuildCobolReport. Slower (parses + writes ~100 rows) and adds sheets to the
' workbook, so it is kept out of Run_All_Tests. Run on demand.
Public Sub Run_RenderTest()
    Test_Begin
    Log_ "Workbook path: " & ThisWorkbook.path
    Test_Phase5.Run_All
    Test_End
End Sub

' Timing harness: prints elapsed ms for each engine stage on ICASE1, with a
' DoEvents after every line so the Immediate window updates even mid-run. The
' last line printed before a freeze identifies the slow stage.
Public Sub Bench_ICASE1()
    Dim cblPath As String, src As String, t As Double
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(cblPath)) = 0 Then Debug.Print "ICASE1 missing: " & cblPath: Exit Sub
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")

    Dim norm As OrderedDict, lines As Collection, root As Collection
    t = Timer: Set norm = CobolParser.Get_NormalizedCobolLines(src, ""): BP_ "normalize", t
    Set lines = norm.Item("Lines")
    t = Timer: Set root = CobolParser.Get_CobolNodes(lines): BP_ "Get_CobolNodes roots=" & root.Count, t

    Dim struct As OrderedDict
    t = Timer: Set struct = CobolParser.Get_ProgramStructure(lines): BP_ "Get_ProgramStructure", t

    CobolParser.ResetEngineState
    Dim e1 As Collection, e2 As Collection, e3 As Collection, e4 As Collection, raw As Collection
    Set e1 = New Collection: Set e2 = New Collection: Set e3 = New Collection: Set e4 = New Collection
    t = Timer: Set raw = CobolParser.Expand_NodeSequence(root, e1, e2, e3, e4)
    BP_ "Expand raw=" & raw.Count & " calls=" & CobolParser.ExpandCalls & " trunc=" & CobolParser.PathTruncated, t

    Dim cg As OrderedDict
    t = Timer: Set cg = CobolParser.Get_CallRelationships(lines, struct): BP_ "Get_CallRelationships edges=" & cg.Item("edges").Count, t

    Dim full As OrderedDict
    t = Timer: Set full = CobolParser.Analyze_Full(src): BP_ "Analyze_Full", t
    Debug.Print "BENCH DONE": DoEvents
End Sub

' Time every stage of the full button flow (analyze -> serialize -> write ->
' parse -> render) so we can see which stage is slow. DoEvents after each line.
Public Sub Bench_Full(Optional ByVal cblName As String = "ICASE2.cbl")
    Dim cblPath As String, src As String, t As Double
    cblPath = ThisWorkbook.path & "\samples\input\" & cblName
    If Len(Dir(cblPath)) = 0 Then Debug.Print "missing: " & cblPath: Exit Sub
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")

    Dim r As OrderedDict
    t = Timer: Set r = CobolParser.Analyze_Full(src, "", "utf-8"): BP_ "Analyze_Full", t

    Dim json As String
    t = Timer: json = JsonWriter.WriteJson(r): BP_ "JsonWriter (len=" & Len(json) & ")", t

    Dim jp As String
    jp = Environ$("TEMP") & "\CobolAnalyzer_bench.logic.json"
    t = Timer: CobolEncoding.WriteAllText jp, json, "utf-8": BP_ "WriteAllText", t

    Dim parsed As Object
    t = Timer: Set parsed = JsonParser.ParseJson(JsonParser.ReadAllText(jp)): BP_ "ReadAllText+ParseJson", t

    t = Timer: CobolLogicViewer.BuildCobolReport jp: BP_ "BuildCobolReport (render)", t

    On Error Resume Next
    Kill jp
    On Error GoTo 0
    Debug.Print "BENCH_FULL DONE": DoEvents
End Sub

Private Sub BP_(ByVal label As String, ByVal t As Double)
    Debug.Print label & ": " & Format$((Timer - t) * 1000, "0") & " ms"
    DoEvents
End Sub

' Run one named test sub via Application.Run, catching any unhandled error
' and converting it to a single [FAIL] line. Lets the rest of the suite run.
Public Sub Run_One(ByVal testName As String)
    On Error Resume Next
    Application.Run testName
    If Err.Number <> 0 Then
        Assert_True False, testName & " threw error #" & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub Log_(ByVal line As String)
    Debug.Print line
    If mLog Is Nothing Then Set mLog = New Collection
    mLog.Add line
End Sub

Private Function SafeToString_(ByVal v As Variant) As String
    On Error Resume Next
    If IsObject(v) Then
        SafeToString_ = "<" & TypeName(v) & ">"
    ElseIf IsNull(v) Then
        SafeToString_ = "<Null>"
    ElseIf IsEmpty(v) Then
        SafeToString_ = "<Empty>"
    Else
        SafeToString_ = CStr(v)
    End If
    On Error GoTo 0
End Function

Private Sub Render_Log()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TestResults")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "TestResults"
    End If
    ws.Cells.Clear
    Dim r As Long, v As Variant
    r = 1
    For Each v In mLog
        ws.Cells(r, 1).value = v
        r = r + 1
    Next v
End Sub
