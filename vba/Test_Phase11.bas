Attribute VB_Name = "Test_Phase11"
' Test_Phase11 - ver3.2: sweep + fallback case generation (Analyze_Flow)
' validated against the ver4 PS oracle on the ICASE3 fixture: 14 arms,
' 3 normal candidate walks (2 seeds + 1 targeted fallback for if-137:else;
' the sweep round gains nothing on ICASE3 because its weight leads away
' from the env value if-132:then needs), 3 abend fallbacks, cases =
' 3 normal + 3 code-abend + 2 synth, full coverage, no infeasible combo.

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_Flow_ICASE3"
    TestRunner.Run_One "Test_Flow_CopyTolerance"
    TestRunner.Run_One "Test_Flow_LoopAndFlag"
End Sub

' (a) loop-form PERFORM (VARYING/UNTIL/TIMES) bodies are inlined once, so
' their branch arms are reachable; (b) the flag idiom (a sibling branch
' MOVEs a literal a later nested IF tests) is covered via the value-driven
' steering retry - the nested arm is reachable by NEITHER seed and the
' sweep cannot fake the flag value, so full coverage REQUIRES steering.
' ver4 oracle expectations: 8 arms, every arm covered, 3 normal cases.
Public Sub Test_Flow_LoopAndFlag()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-IN    PIC X(01)." & vbLf
    s = s & "       01  W-X     PIC X(01)." & vbLf
    s = s & "       01  F-MAS1  PIC X(01)." & vbLf
    s = s & "       01  T-VAL   PIC 9(01)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       01  I-IDX   PIC 9(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           PERFORM SETUP-SEC." & vbLf
    s = s & "           PERFORM LOOP-SEC VARYING I-IDX FROM 1 BY 1" & vbLf
    s = s & "               UNTIL I-IDX > 20." & vbLf
    s = s & "           IF W-X = '9'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               IF F-MAS1 = 'B'" & vbLf
    s = s & "               THEN" & vbLf
    s = s & "                   MOVE 'BB' TO W-OUT" & vbLf
    s = s & "               ELSE" & vbLf
    s = s & "                   MOVE 'KK' TO W-OUT" & vbLf
    s = s & "               END-IF" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'ZZ' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       SETUP-SEC SECTION." & vbLf
    s = s & "       SETUP-000." & vbLf
    s = s & "           IF W-IN = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'B' TO F-MAS1" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'K' TO F-MAS1" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       SETUP-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       LOOP-SEC SECTION." & vbLf
    s = s & "       LOOP-000." & vbLf
    s = s & "           IF T-VAL = 0" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'X1' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'Y1' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       LOOP-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(8), CLng(flow.Item("arms").Count), "loop+flag: 8 arms"

    Dim a As OrderedDict, c As OrderedDict, v As Variant, covered As Boolean, uncov As Long
    For Each a In flow.Item("arms")
        covered = False
        For Each c In flow.Item("cases")
            For Each v In c.Item("arms")
                If CStr(v) = CStr(a.Item("Token")) Then covered = True
            Next v
        Next c
        If Not covered Then uncov = uncov + 1
    Next a
    TestRunner.Assert_Equal CLng(0), uncov, _
        "loop bodies inlined + flag idiom steered: every arm covered"
    ' the nested flag arm is reachable by NEITHER seed (the ELSE-pref seed
    ' turns away at W-X) and the sweep cannot fake the flag value, so
    ' uncovered=0 above REQUIRES the value-steering fallback walk.
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("cases").Count), "loop+flag: 3 cases"
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("normalPaths")), _
        "2 seeds + 1 steered fallback walk"
End Sub

' Regression: real sources are full of COPY ... PREFIXING lines, whose data
' items carry an EMPTY level. CLng("") on that level crashed Analyze_Flow
' (type mismatch) on every COPY-bearing program while the COPY-less fixture
' passed. Analyze_Flow must tolerate them.
Public Sub Test_Flow_CopyTolerance()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-PARAM." & vbLf
    s = s & "           COPY DUMMYCPY PREFIXING W-." & vbLf
    s = s & "       01  W-FLG  PIC X(01)." & vbLf
    s = s & "       01  W-OUT  PIC X(01)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           IF W-FLG = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'A' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           GOBACK." & vbLf

    Dim terms As Collection
    Set terms = New Collection
    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, terms)
    TestRunner.Assert_Equal CLng(2), CLng(flow.Item("arms").Count), "COPY-bearing source: 2 arms"
    TestRunner.Assert_True flow.Item("cases").Count >= 2, "COPY-bearing source: cases generated (no crash)"
End Sub

Public Sub Test_Flow_ICASE3()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE3.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE3.cbl missing: " & p
        Exit Sub
    End If
    Dim src As String
    src = CobolEncoding.ReadCobolSource(p, "auto")

    Dim terms As Collection
    Set terms = New Collection
    terms.Add "S99-ABEND-PROC"

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(src, terms)

    TestRunner.Assert_Equal CLng(14), CLng(flow.Item("arms").Count), "ICASE3 arms = 14"
    ' 2 seeds + 1 targeted fallback (the ICASE3 sweep round gains 0)
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("normalPaths")), "normal candidate walks = 3"
    TestRunner.Assert_True Not CBool(flow.Item("truncated")), "not truncated"

    Dim cases As Collection
    Set cases = flow.Item("cases")
    TestRunner.Assert_Equal CLng(8), CLng(cases.Count), "cases = 8"

    Dim c As OrderedDict, nN As Long, nA As Long, nS As Long
    For Each c In cases
        Select Case CStr(c.Item("kind"))
            Case "normal": nN = nN + 1
            Case "abend":  nA = nA + 1
            Case "synth":  nS = nS + 1
        End Select
    Next c
    TestRunner.Assert_Equal CLng(3), nN, "normal cases = 3"
    TestRunner.Assert_Equal CLng(3), nA, "code-derived abend cases = 3"
    TestRunner.Assert_Equal CLng(2), nS, "synthesized call-failure cases = 2"

    ' full arm coverage across all cases
    Dim a As OrderedDict, covered As Boolean, uncovered As Long, v As Variant
    For Each a In flow.Item("arms")
        covered = False
        For Each c In cases
            For Each v In c.Item("arms")
                If CStr(v) = CStr(a.Item("Token")) Then covered = True
            Next v
        Next c
        If Not covered Then uncovered = uncovered + 1
    Next a
    TestRunner.Assert_Equal CLng(0), uncovered, "coverage matrix fully covered"

    ' dataflow pruning: result-flag contradictions never co-occur
    Dim bad As Boolean
    For Each c In cases
        If HasArm_(c, "if-180:then") And HasArm_(c, "if-132:then") Then bad = True
        If HasArm_(c, "if-176:then") And HasArm_(c, "if-132:else") Then bad = True
    Next c
    TestRunner.Assert_True Not bad, "no infeasible arm combination (constant propagation)"

    ' synthesized targets: DATESUB + SUBX, never ABSUB (inside terminator)
    Dim hasDate As Boolean, hasSubx As Boolean, hasAbsub As Boolean
    For Each c In cases
        If CStr(c.Item("term")) = "synth:DATESUB" Then hasDate = True
        If CStr(c.Item("term")) = "synth:SUBX" Then hasSubx = True
        If CStr(c.Item("term")) = "synth:ABSUB" Then hasAbsub = True
    Next c
    TestRunner.Assert_True hasDate And hasSubx, "synth cases for DATESUB and SUBX"
    TestRunner.Assert_True Not hasAbsub, "no synth case for ABSUB (terminator section)"

    ' the business-sub THEN case ends at the PA400 assignment (final action)
    For Each c In cases
        If CStr(c.Item("kind")) = "normal" And HasArm_(c, "if-137:then") Then
            TestRunner.Assert_Equal CLng(139), CLng(c.Item("finalLine")), _
                "final action of the sub-return case = MOVE SBX-PA400 (line 139)"
            TestRunner.Assert_Equal "goback", CStr(c.Item("term")), "normal case terminates at GOBACK"
        End If
    Next c

    ' terminator auto-detection: *ABEND* section names need no registration
    Dim flowAuto As OrderedDict
    Set flowAuto = CobolFlow.Analyze_Flow(src, New Collection)
    TestRunner.Assert_Equal CLng(8), CLng(flowAuto.Item("cases").Count), _
        "ABEND-named terminator auto-detected without registration (8 cases)"
End Sub

Private Function HasArm_(ByVal c As OrderedDict, ByVal token As String) As Boolean
    Dim v As Variant
    HasArm_ = False
    For Each v In c.Item("arms")
        If CStr(v) = token Then HasArm_ = True
    Next v
End Function
