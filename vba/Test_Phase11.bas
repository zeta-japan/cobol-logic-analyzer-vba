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
    TestRunner.Run_One "Test_Flow_UnsetSteer"
    TestRunner.Run_One "Test_Flow_ReadAheadIf"
    TestRunner.Run_One "Test_Flow_ReadAheadEval"
    TestRunner.Run_One "Test_Flow_ArithAndGoto"
    TestRunner.Run_One "Test_Flow_CallArgUnset"
    TestRunner.Run_One "Test_Flow_BlockerSteer"
    TestRunner.Run_One "Test_Flow_BlockerSteerEval"
    TestRunner.Run_One "Test_Flow_ArmMeta"
End Sub

' ver3.3 deliverable plumbing on the BlockerSteer fixture: arms carry
' MarkLine (first statement INSIDE the arm - where Besshi-1 marks the TC),
' the result exposes the SECTION ranges, and CobolXdm.BuildTcMap maps
' every covered arm to the FIRST covering case.
' Fixture lines: if-14 (W-EOF gate): then arm = CONTINUE only (no AST
' node -> MarkLine falls back to the branch line 14), else-child
' PERFORM=18; if-20 (F-MAS1): then-child inner if=22 (else arm is
' CONTINUE-only -> falls back to 20, not pinned);
' if-22 (compound): then MOVE=24, else MOVE=26. Sections: MAIN-PROC=11,
' READ-SEC=32.
Public Sub Test_Flow_ArmMeta()
    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(BlockerSteerSrc_(), New Collection)

    Dim a As OrderedDict, mk As OrderedDict
    Set mk = New OrderedDict
    For Each a In flow.Item("arms")
        mk.Add CStr(a.Item("Token")), CLng(a.Item("MarkLine"))
    Next a
    TestRunner.Assert_Equal CLng(14), CLng(mk.Item("if-14:then")), "CONTINUE-only arm marks at the branch line (no AST node)"
    TestRunner.Assert_Equal CLng(18), CLng(mk.Item("if-14:else")), "else arm marks at first inner statement"
    TestRunner.Assert_Equal CLng(22), CLng(mk.Item("if-20:then")), "nested-if arm marks at the inner IF line"
    TestRunner.Assert_Equal CLng(26), CLng(mk.Item("if-22:else")), "compound-cond else marks at its MOVE"

    Dim secs As Collection
    Set secs = flow.Item("sections")
    TestRunner.Assert_Equal CLng(2), CLng(secs.Count), "two sections exported"
    TestRunner.Assert_Equal "MAIN-PROC", CStr(secs(1).Item("name")), "first section name"
    TestRunner.Assert_Equal CLng(11), CLng(secs(1).Item("line")), "first section start line"

    Dim map As OrderedDict, v As Variant, ok As Boolean
    Set map = CobolXdm.BuildTcMap(flow)
    For Each a In flow.Item("arms")
        TestRunner.Assert_True map.Exists(CStr(a.Item("Token"))), "every covered arm mapped: " & CStr(a.Item("Token"))
        ' consistency: the mapped case really contains the token
        Dim c As OrderedDict
        ok = False
        For Each c In flow.Item("cases")
            If CStr(c.Item("id")) = CStr(map.Item(CStr(a.Item("Token")))) Then
                For Each v In c.Item("arms")
                    If CStr(v) = CStr(a.Item("Token")) Then ok = True
                Next v
            End If
        Next c
        TestRunner.Assert_True ok, "mapped case covers the token: " & CStr(a.Item("Token"))
    Next a
End Sub

' the production shape behind the compound-condition red rows: a
' IF nested under EVALUATE WHEN 'D', flag initialized to space, only
' setter a GATED group MOVE. Fallback walks for the inner arms conflict
' AT the EVALUATE (mMissTok = when-D id), the target has no steer of its
' own, so coverage needs the blocker-keyed retry + descendant unset.
' The oracle has no EVALUATE parser - pins derived by hand: arms = 7
' (eof-if 2, WHENs 3, inner-if 2), all covered.
Public Sub Test_Flow_BlockerSteerEval()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  FD-REC  PIC X(02)." & vbLf
    s = s & "       01  W-WK." & vbLf
    s = s & "           03  F-MAS1  PIC X(01)." & vbLf
    s = s & "           03  F-OTH   PIC X(01)." & vbLf
    s = s & "       01  DT1     PIC 9(03)." & vbLf
    s = s & "       01  DT2     PIC 9(03)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       01  W-EOF   PIC X(01)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ' ' TO F-MAS1." & vbLf
    s = s & "           IF W-EOF = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               PERFORM READ-SEC" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           EVALUATE F-MAS1" & vbLf
    s = s & "               WHEN ' '" & vbLf
    s = s & "                   CONTINUE" & vbLf
    s = s & "               WHEN 'D'" & vbLf
    s = s & "                   IF DT1 NOT = 0 AND DT2 NOT = 0" & vbLf
    s = s & "                   THEN" & vbLf
    s = s & "                       MOVE 'AA' TO W-OUT" & vbLf
    s = s & "                   ELSE" & vbLf
    s = s & "                       MOVE 'BB' TO W-OUT" & vbLf
    s = s & "                   END-IF" & vbLf
    s = s & "               WHEN OTHER" & vbLf
    s = s & "                   MOVE 'ZZ' TO W-OUT" & vbLf
    s = s & "           END-EVALUATE." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       READ-SEC SECTION." & vbLf
    s = s & "       READS-000." & vbLf
    s = s & "           MOVE FD-REC TO W-WK." & vbLf
    s = s & "       READS-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(7), CLng(flow.Item("arms").Count), "eval blocker-steer: 7 arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "compound-cond arms under a blocked WHEN: all covered"
End Sub

' blocker-keyed retry + group-MOVE descendant unset: the target arm has a
' COMPOUND condition (no steer info of its own); the conflict happens at
' the ANCESTOR IF on F-MAS1, whose setter is a GROUP move GATED behind a
' sibling branch (so the propagated constant survives attempt 1 - without
' the gate the unconditional Invalidate_ would mask the retry machinery),
' i.e. MOVE FD-REC TO
' W-WK with F-MAS1 inside W-WK). The retry must key on the blocking arm
' and havoc F-MAS1 via the descendant-expanded unset site.
' ver4 oracle (BIGCASE7): 6 arms, all covered, 4 normal paths, 3 cases.
Public Sub Test_Flow_BlockerSteer()
    Dim s As String
    s = BlockerSteerSrc_()

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(6), CLng(flow.Item("arms").Count), "blocker-steer: 6 arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "compound-cond arms under a blocked ancestor: all covered"
    TestRunner.Assert_Equal CLng(4), CLng(flow.Item("normalPaths")), "blocker-steer: 4 normal paths"
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("cases").Count), "blocker-steer: 3 cases"
End Sub

' a CALL's USING args may be modified by the callee (ADABAS RC idiom):
' the call site registers them as unset sites, so a later test blocked by
' the propagated init constant is routed through the calling branch.
' ver4 oracle (BIGCASE6): the CALL sits on the NON-preferred arm and
' passes a GROUP param while the test reads a subordinate field - the
' arm is reachable ONLY via the call-site unset chain (incl. descendant
' expansion). 4 arms, all covered, 3 normal paths, 3 cases.
Public Sub Test_Flow_CallArgUnset()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-SW    PIC X(01)." & vbLf
    s = s & "       01  W-PARAM." & vbLf
    s = s & "           03  W-RC    PIC 9(01)." & vbLf
    s = s & "           03  W-DATA  PIC X(10)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ZERO TO W-RC." & vbLf
    s = s & "           IF W-SW = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               PERFORM CALL-SEC" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           IF W-RC = 3" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'AA' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'BB' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       CALL-SEC SECTION." & vbLf
    s = s & "       CALLS-000." & vbLf
    s = s & "           CALL 'SUBX' USING W-PARAM." & vbLf
    s = s & "       CALLS-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(4), CLng(flow.Item("arms").Count), "call-arg: 4 arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "RC=3 arm covered by routing through the CALL branch"
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("normalPaths")), "call-arg: 2 seeds + 1 unset-steered retry"
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("cases").Count), "call-arg: 2 normal + 1 synth"
End Sub

' (a) arithmetic verbs invalidate their targets - a MOVE ZERO + ADD 1
' counter must not keep the stale constant (it falsely pruned both arms
' of later tests); (b) GO TO <name> is followed as a forward exit-jump
' (Natural ESCAPE BOTTOM conversion style), so arms beyond/inside the
' jump target are reachable; default arms avoid jump-away branches.
' ver4 oracle (BIGCASE5): 6 arms, all covered, 4 cases.
Public Sub Test_Flow_ArithAndGoto()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  C-NUM   PIC 9(03)." & vbLf
    s = s & "       01  W-ERR   PIC X(01)." & vbLf
    s = s & "       01  W-FIN   PIC X(01)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ZERO TO C-NUM." & vbLf
    s = s & "           PERFORM PROC-SEC." & vbLf
    s = s & "           PERFORM TAIL-SEC." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       PROC-SEC SECTION." & vbLf
    s = s & "       PROC-000." & vbLf
    s = s & "           ADD 1 TO C-NUM." & vbLf
    s = s & "           IF W-ERR = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               GO TO EXIT-SEC" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           IF C-NUM NOT = 0" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'CC' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'DD' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       PROC-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       EXIT-SEC SECTION." & vbLf
    s = s & "       EXITS-000." & vbLf
    s = s & "           IF W-FIN = '9'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'EE' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'FF' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       EXITS-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       TAIL-SEC SECTION." & vbLf
    s = s & "       TAIL-000." & vbLf
    s = s & "           MOVE 'TT' TO W-OUT." & vbLf
    s = s & "       TAIL-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(6), CLng(flow.Item("arms").Count), "arith+goto: 6 arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "counter arms unpruned + arms beyond GO TO reachable: all covered"
    TestRunner.Assert_Equal CLng(4), CLng(flow.Item("cases").Count), "arith+goto: 4 cases"
End Sub

' read-ahead loop idiom: the flag is initialized to a literal and the real
' setter (non-literal MOVE) runs at the BOTTOM of the loop body, i.e. AFTER
' the test on our single inlined pass. Value/unset steering cannot reorder
' the walk, so the havoc retry treats the flag as unknown. ver4 oracle:
' 2 arms, all covered, 2 cases.
Public Sub Test_Flow_ReadAheadIf()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  F-EOF   PIC X(01)." & vbLf
    s = s & "       01  FD-KBN  PIC X(01)." & vbLf
    s = s & "       01  F-MAS1  PIC X(01)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ' ' TO F-MAS1." & vbLf
    s = s & "           PERFORM LOOP-SEC UNTIL F-EOF = '1'." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       LOOP-SEC SECTION." & vbLf
    s = s & "       LOOP-000." & vbLf
    s = s & "           IF F-MAS1 = 'B'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'BB' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'KK' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           MOVE FD-KBN TO F-MAS1." & vbLf
    s = s & "       LOOP-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(2), CLng(flow.Item("arms").Count), "read-ahead IF: 2 arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "loop-carried flag arm covered via havoc retry"
    TestRunner.Assert_Equal CLng(2), CLng(flow.Item("cases").Count), "read-ahead IF: 2 cases"
End Sub

' same idiom with EVALUATE (the shape seen in real sources): WHEN literals
' plus WHEN OTHER, flag set from a record field after the EVALUATE.
Public Sub Test_Flow_ReadAheadEval()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  F-EOF   PIC X(01)." & vbLf
    s = s & "       01  FD-KBN  PIC X(01)." & vbLf
    s = s & "       01  F-MAS1  PIC X(01)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ' ' TO F-MAS1." & vbLf
    s = s & "           PERFORM LOOP-SEC UNTIL F-EOF = '1'." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       LOOP-SEC SECTION." & vbLf
    s = s & "       LOOP-000." & vbLf
    s = s & "           EVALUATE F-MAS1" & vbLf
    s = s & "               WHEN ' '" & vbLf
    s = s & "                   CONTINUE" & vbLf
    s = s & "               WHEN 'B'" & vbLf
    s = s & "                   MOVE 'BB' TO W-OUT" & vbLf
    s = s & "               WHEN OTHER" & vbLf
    s = s & "                   MOVE 'ZZ' TO W-OUT" & vbLf
    s = s & "           END-EVALUATE." & vbLf
    s = s & "           MOVE FD-KBN TO F-MAS1." & vbLf
    s = s & "       LOOP-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("arms").Count), "read-ahead EVALUATE: 3 arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "WHEN 'B' and WHEN OTHER covered via havoc retry"
    TestRunner.Assert_Equal CLng(3), CLng(flow.Item("cases").Count), "read-ahead EVALUATE: 3 cases"
End Sub

Private Function UncovCount_(ByVal flow As OrderedDict) As Long
    Dim a As OrderedDict, c As OrderedDict, v As Variant, covered As Boolean, n As Long
    For Each a In flow.Item("arms")
        covered = False
        For Each c In flow.Item("cases")
            For Each v In c.Item("arms")
                If CStr(v) = CStr(a.Item("Token")) Then covered = True
            Next v
        Next c
        If Not covered Then n = n + 1
    Next a
    UncovCount_ = n
End Function

' DB-flag idiom: F-MAS1 is initialized to a literal and set from a RECORD
' FIELD (non-literal MOVE) in a sibling branch - no literal 'B' site exists,
' so value steering alone cannot unblock IF F-MAS1 = 'B'. The unset-steering
' fallback walks through the invalidating MOVE instead. ver4 oracle: 4 arms,
' all covered, 2 cases, 3 normal candidate walks.
Public Sub Test_Flow_UnsetSteer()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-EOF   PIC X(01)." & vbLf
    s = s & "       01  FD-KBN  PIC X(01)." & vbLf
    s = s & "       01  F-MAS1  PIC X(01)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ' ' TO F-MAS1." & vbLf
    s = s & "           IF W-EOF = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE FD-KBN TO F-MAS1" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           IF F-MAS1 = 'B'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'BB' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'KK' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           GOBACK." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(4), CLng(flow.Item("arms").Count), "unset-steer: 4 arms"

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
        "DB-flag arm covered via unset steering (no literal 'B' site)"
    TestRunner.Assert_Equal CLng(2), CLng(flow.Item("cases").Count), "unset-steer: 2 cases"
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

' shared fixture for BlockerSteer / ArmMeta (line numbers pinned)
Private Function BlockerSteerSrc_() As String
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  FD-REC  PIC X(02)." & vbLf
    s = s & "       01  W-WK." & vbLf
    s = s & "           03  F-MAS1  PIC X(01)." & vbLf
    s = s & "           03  F-OTH   PIC X(01)." & vbLf
    s = s & "       01  DT1     PIC 9(03)." & vbLf
    s = s & "       01  DT2     PIC 9(03)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       01  W-EOF   PIC X(01)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           MOVE ' ' TO F-MAS1." & vbLf
    s = s & "           IF W-EOF = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               PERFORM READ-SEC" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           IF F-MAS1 = 'D'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               IF DT1 NOT = 0 AND DT2 NOT = 0" & vbLf
    s = s & "               THEN" & vbLf
    s = s & "                   MOVE 'AA' TO W-OUT" & vbLf
    s = s & "               ELSE" & vbLf
    s = s & "                   MOVE 'BB' TO W-OUT" & vbLf
    s = s & "               END-IF" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       READ-SEC SECTION." & vbLf
    s = s & "       READS-000." & vbLf
    s = s & "           MOVE FD-REC TO W-WK." & vbLf
    s = s & "       READS-999." & vbLf
    s = s & "           EXIT." & vbLf
    BlockerSteerSrc_ = s
End Function
