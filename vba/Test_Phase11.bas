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
    TestRunner.Run_One "Test_Flow_DeepAbendNormalCover"
    TestRunner.Run_One "Test_Flow_TransitiveAbend"
    TestRunner.Run_One "Test_Flow_BlockerSteer"
    TestRunner.Run_One "Test_Flow_BlockerSteerEval"
    TestRunner.Run_One "Test_Flow_ArmMeta"
    TestRunner.Run_One "Test_Xdm_CondJp"
    TestRunner.Run_One "Test_Xdm_ActionJp"
    TestRunner.Run_One "Test_CaseView_DeadSection"
    TestRunner.Run_One "Test_Flow_OrphanEntry"
    TestRunner.Run_One "Test_Flow_FallThroughEntry"
End Sub

' DeadSection_ classifies an uncovered arm as a confidently-dead SECTION
' (owning section known, NO PERFORM caller, NO textual reference) so the
' coverage sheets show it out-of-scope (gray, excluded from the denominator)
' instead of uncovered (red). The "callers>0 = analysis gap" and "refs>0"
' sub-cases must NOT be classified dead - they stay red/counted as possible
' tool gaps to report.
Public Sub Test_CaseView_DeadSection()
    Dim flow As OrderedDict, ad As OrderedDict, sec As String
    Set flow = New OrderedDict
    Set ad = New OrderedDict
    ad.Add "t_dead", "noctx|S020-UPDATE||"
    ad.Add "t_gap", "noctx|S020-UPDATE|CALLER-SEC|"
    ad.Add "t_ref", "noctx|S020-UPDATE||L100:GO TO S020"
    ad.Add "t_nosec", "noctx|||"
    ad.Add "t_conflict", "conflict|W-X=ZERO|tried"
    flow.Add "armDiag", ad

    sec = "x"
    TestRunner.Assert_True CobolCaseView.DeadSection_(flow, "t_dead", sec), "no caller + no ref -> dead"
    TestRunner.Assert_Equal "S020-UPDATE", sec, "dead case returns the owning section name"
    TestRunner.Assert_True Not CobolCaseView.DeadSection_(flow, "t_gap", sec), "callers>0 (analysis gap) is NOT dead"
    TestRunner.Assert_True Not CobolCaseView.DeadSection_(flow, "t_ref", sec), "refs>0 (textual ref) is NOT dead"
    TestRunner.Assert_True Not CobolCaseView.DeadSection_(flow, "t_nosec", sec), "empty section is NOT dead"
    TestRunner.Assert_True Not CobolCaseView.DeadSection_(flow, "t_conflict", sec), "conflict diag is NOT dead"
    TestRunner.Assert_True Not CobolCaseView.DeadSection_(flow, "t_missing", sec), "unknown token is NOT dead"

    ' secOut is cleared on every call (set to "" before any classification)
    Dim ok As Boolean
    ok = CobolCaseView.DeadSection_(flow, "t_gap", sec)
    TestRunner.Assert_Equal "", sec, "secOut cleared for a non-dead arm"
End Sub

' ver3.9 unit-test entries: DEAD-UPD-SEC is never PERFORMed (an orphan), but a
' unit driver can PERFORM it directly, so its EVALUATE arms must get cases. The
' main flow only reaches WORK-SEC; the 3 WHEN arms of DEAD-UPD-SEC are coverable
' ONLY via the orphan-entry phase, so full coverage proves that phase ran. The
' WHEN OTHER arm PERFORMs the ABEND section, exercising the orphan abend path.
Public Sub Test_Flow_OrphanEntry()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-FLAG  PIC X(01)." & vbLf
    s = s & "       01  W-X     PIC X(02)." & vbLf
    s = s & "       01  W-RC    PIC X(01)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           PERFORM WORK-SEC." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       WORK-SEC SECTION." & vbLf
    s = s & "       WORK-000." & vbLf
    s = s & "           IF W-FLAG = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE '11' TO W-X" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE '22' TO W-X" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       WORK-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       DEAD-UPD-SEC SECTION." & vbLf
    s = s & "       DEAD-000." & vbLf
    s = s & "           CALL 'SUBR' USING W-RC." & vbLf
    s = s & "           EVALUATE W-RC" & vbLf
    s = s & "             WHEN ZERO" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "             WHEN '1'" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "             WHEN OTHER" & vbLf
    s = s & "               PERFORM Z-ABEND-SEC" & vbLf
    s = s & "           END-EVALUATE." & vbLf
    s = s & "       DEAD-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       Z-ABEND-SEC SECTION." & vbLf
    s = s & "       Z-000." & vbLf
    s = s & "           DISPLAY 'ABEND'." & vbLf
    s = s & "           GOBACK." & vbLf

    Dim flow As OrderedDict, noTerms As Collection
    Set noTerms = New Collection
    Set flow = CobolFlow.Analyze_Flow(s, noTerms)

    TestRunner.Assert_True flow.Item("arms").Count >= 5, _
        "orphan fixture exposes WORK IF (2) + DEAD-UPD EVALUATE (3) arms: " & flow.Item("arms").Count
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "orphan-entry: every arm covered, incl. the never-PERFORMed DEAD-UPD-SEC"

    ' prove the coverage came from entering DEAD-UPD-SEC directly (unit entry)
    Dim c As OrderedDict, e As OrderedDict, sawOrphan As Boolean
    For Each c In flow.Item("cases")
        For Each e In c.Item("events")
            If InStr(CStr(e.Item("Text")), "DEAD-UPD-SEC") > 0 Then sawOrphan = True
        Next e
    Next c
    TestRunner.Assert_True sawOrphan, "a generated case enters DEAD-UPD-SEC as a unit-test entry"
End Sub

' a main entry that ends with a bare EXIT (no GOBACK/STOP RUN) - a common
' Natural-to-COBOL idiom - falls off the procedure division = implicit normal
' return. Without the RunWalk_ promotion every normal walk has Term="" and is
' dropped, leaving every arm flagged as a dead path (path-construction-
' impossible). Here both IF arms must be covered by normal cases.
Public Sub Test_Flow_FallThroughEntry()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-FLAG  PIC X(01)." & vbLf
    s = s & "       01  W-X     PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           IF W-FLAG = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE '11' TO W-X" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE '22' TO W-X" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       MAIN-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict, noTerms As Collection
    Set noTerms = New Collection
    Set flow = CobolFlow.Analyze_Flow(s, noTerms)

    TestRunner.Assert_Equal CLng(2), CLng(flow.Item("arms").Count), "fall-through entry: 2 IF arms"
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), _
        "bare-EXIT entry: both arms covered (end-of-program = implicit normal return)"

    Dim c As OrderedDict, hasNormal As Boolean
    For Each c In flow.Item("cases")
        If CStr(c.Item("kind")) = "normal" Then hasNormal = True
    Next c
    TestRunner.Assert_True hasNormal, "a NORMAL case is generated though there is no explicit GOBACK"
End Sub

' the pattern draft now lists straight-line statements too; ActionJp_
' routes each verb to its template. ASCII-structural checks (identifiers
' are preserved, terminators are skipped) - the JP wording is by eye.
Public Sub Test_Xdm_ActionJp()
    TestRunner.Assert_Equal "", CobolXdm.ActionJpOf("EXIT"), "bare EXIT is skipped (empty)"
    TestRunner.Assert_Equal "", CobolXdm.ActionJpOf("CONTINUE"), "CONTINUE is skipped (empty)"

    Dim r As String
    r = CobolXdm.ActionJpOf("ADD 1 TO CNT-FI1")
    TestRunner.Assert_True InStr(r, "CNT-FI1") > 0 And InStr(r, "1") > 0, "ADD keeps operand + target"
    TestRunner.Assert_Equal "", CobolXdm.ActionJpOf("MOVE WK-A TO WK-B"), "MOVE with no output records -> skipped"
    TestRunner.Assert_Equal "", CobolXdm.ActionJpOf2("MOVE WK-A TO WK-SUB-PARM", "N-PARAM"), "MOVE to a non-output item -> skipped"
    TestRunner.Assert_True InStr(CobolXdm.ActionJpOf2("MOVE WS-X TO N-PARAM-PA310", "N-PARAM"), "N-PARAM-PA310") > 0, "MOVE to an output (LINKAGE) item -> kept"
    TestRunner.Assert_Equal "", CobolXdm.ActionJpOf("INITIALIZE F-SEL1"), "INITIALIZE is housekeeping (skipped)"
    TestRunner.Assert_Equal "", CobolXdm.ActionJpOf("GOBACK"), "GOBACK is housekeeping (skipped)"

    ' loop form: the UNTIL condition operands survive (loop routing)
    r = CobolXdm.ActionJpOf("PERFORM UNTIL DCP-WDCPRC1 = DCP-EOF")
    TestRunner.Assert_True InStr(r, "DCP-WDCPRC1") > 0 And InStr(r, "DCP-EOF") > 0, "PERFORM UNTIL keeps the condition"

    ' simple perform with no comment falls back to the label
    r = CobolXdm.ActionJpOf("PERFORM S010-FI1-READ-PROC")
    TestRunner.Assert_True InStr(r, "S010-FI1-READ-PROC") > 0, "simple PERFORM keeps the target"
    r = CobolXdm.ActionJpOf("SUBTRACT WK-X FROM WK-Y")
    TestRunner.Assert_True InStr(r, "WK-X") > 0 And InStr(r, "WK-Y") > 0, "SUBTRACT keeps operand + target"
    r = CobolXdm.ActionJpOf("GO TO ERR-EXIT")
    TestRunner.Assert_True InStr(r, "ERR-EXIT") > 0, "GO TO keeps the target"
End Sub

' template-JP condition translation: operators replaced, identifiers
' and literals preserved (the JP words themselves are checked by eye -
' this file is ASCII, so assertions are structural)
Public Sub Test_Xdm_CondJp()
    Dim r As String
    r = CobolXdm.CondJp_("W-PA210 = '97' OR W-PB210 = '99000010'")
    TestRunner.Assert_True InStr(r, " OR ") = 0, "OR translated away"
    TestRunner.Assert_True InStr(r, "'97'") > 0, "literal preserved"
    TestRunner.Assert_True InStr(r, "W-PA210") = 1, "item name preserved at head"
    r = CobolXdm.CondJp_("DT1 NOT = 0 AND DT2 NOT = 0")
    TestRunner.Assert_True InStr(r, " AND ") = 0, "AND translated away"
    TestRunner.Assert_True InStr(r, "NOT") = 0, "NOT= folded into one operator"
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

' B fix (ver3.5): a loop-index boundary arm (IF I-IDX1 NOT = 26 ELSE,
' i.e. I-IDX1 = 26) is normally reachable, but the ELSE-preferring seed
' descends into a DEEP abend (PERFORM MID-SEC -> U-ABEND, beyond the
' shallow abend-avoidance) and "claims" the arm, leaving it covered only
' by the ?? case. The second normal-coverage pass must give it a normal
' case. Mirrors PGM-N if-238:else. Oracle BIGCASE9: normal=3, abend=1.
Public Sub Test_Flow_DeepAbendNormalCover()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-CH    PIC X(01)." & vbLf
    s = s & "       01  W-CA    PIC X(01)." & vbLf
    s = s & "       01  W-RC    PIC X(01)." & vbLf
    s = s & "       01  I-IDX1  PIC 9(02)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           PERFORM LOOP-SEC." & vbLf
    s = s & "           PERFORM CALL-SEC." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       LOOP-SEC SECTION." & vbLf
    s = s & "       LOOP-000." & vbLf
    s = s & "           PERFORM VARYING I-IDX1 FROM 1 BY 1" & vbLf
    s = s & "               UNTIL I-IDX1 > 26" & vbLf
    s = s & "               IF W-CH = W-CA" & vbLf
    s = s & "               THEN" & vbLf
    s = s & "                   MOVE 'AA' TO W-OUT" & vbLf
    s = s & "               ELSE" & vbLf
    s = s & "                   IF I-IDX1 NOT = 26" & vbLf
    s = s & "                   THEN" & vbLf
    s = s & "                       MOVE 'BB' TO W-OUT" & vbLf
    s = s & "                   ELSE" & vbLf
    s = s & "                       CONTINUE" & vbLf
    s = s & "                   END-IF" & vbLf
    s = s & "               END-IF" & vbLf
    s = s & "           END-PERFORM." & vbLf
    s = s & "       LOOP-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       CALL-SEC SECTION." & vbLf
    s = s & "       CALL-000." & vbLf
    s = s & "           IF W-RC = '0'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'OK' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               PERFORM MID-SEC" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       CALL-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       MID-SEC SECTION." & vbLf
    s = s & "       MID-000." & vbLf
    s = s & "           PERFORM U-ABEND-PROC." & vbLf
    s = s & "       MID-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       U-ABEND-PROC SECTION." & vbLf
    s = s & "       UAB-000." & vbLf
    s = s & "           CALL 'ERRSUB'." & vbLf
    s = s & "       UAB-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), "deep-abend fixture fully covered"
    Dim tok As String
    tok = ArmTokenByDisp_(flow, "IF I-IDX1 NOT = 26 [ELSE]")
    TestRunner.Assert_True Len(tok) > 0, "found the idx-equality boundary arm"
    TestRunner.Assert_True NormCovers_(flow, tok), "idx-equality boundary arm covered by a NORMAL case (not only abnormal)"
End Sub

' token of the arm whose display text matches (or "" if none)
Private Function ArmTokenByDisp_(ByVal flow As OrderedDict, ByVal disp As String) As String
    ArmTokenByDisp_ = ""
    Dim a As OrderedDict
    For Each a In flow.Item("arms")
        If CStr(a.Item("Disp")) = disp Then
            ArmTokenByDisp_ = CStr(a.Item("Token"))
            Exit Function
        End If
    Next a
End Function

' True if a NORMAL-kind case covers the arm token
Private Function NormCovers_(ByVal flow As OrderedDict, ByVal token As String) As Boolean
    NormCovers_ = False
    Dim c As OrderedDict, v As Variant
    For Each c In flow.Item("cases")
        If CStr(c.Item("kind")) = "normal" Then
            For Each v In c.Item("arms")
                If CStr(v) = token Then
                    NormCovers_ = True
                    Exit Function
                End If
            Next v
        End If
    Next c
End Function

' transitive ABEND (ver3.7): ERRHND-SEC unconditionally PERFORMs the abend
' section but is NOT named *ABEND*. It is PERFORM'd from CHECK-SEC which is
' PERFORM'd early in PRE-SEC. Without transitive-terminator detection the
' default walk takes CHECK's handler arm, dies, and blocks every downstream
' arm (PRE's else, TAIL's then). Transitivity marks ERRHND-SEC a terminator
' so abend-avoidance steers around it. Mirrors the real-program always-abend error-handler pattern.
' Oracle BIGCASE10: 4 cases (2 normal + 2 abend), uncovered 0.
Public Sub Test_Flow_TransitiveAbend()
    Dim s As String
    s = ""
    s = s & "       WORKING-STORAGE SECTION." & vbLf
    s = s & "       01  W-FLG   PIC X(01)." & vbLf
    s = s & "       01  W-RC1   PIC X(01)." & vbLf
    s = s & "       01  W-W     PIC X(01)." & vbLf
    s = s & "       01  W-OUT   PIC X(02)." & vbLf
    s = s & "       PROCEDURE DIVISION." & vbLf
    s = s & "       MAIN-PROC SECTION." & vbLf
    s = s & "       MAIN-000." & vbLf
    s = s & "           PERFORM PRE-SEC." & vbLf
    s = s & "           PERFORM TAIL-SEC." & vbLf
    s = s & "           GOBACK." & vbLf
    s = s & "       PRE-SEC SECTION." & vbLf
    s = s & "       PRE-000." & vbLf
    s = s & "           PERFORM CHECK-SEC." & vbLf
    s = s & "           IF W-RC1 = '0'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'OK' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 999 TO W-OUT" & vbLf
    s = s & "               PERFORM U-ABEND-PROC" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       PRE-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       CHECK-SEC SECTION." & vbLf
    s = s & "       CHECK-000." & vbLf
    s = s & "           IF W-FLG = '1'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               PERFORM ERRHND-SEC" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               CONTINUE" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       CHECK-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       ERRHND-SEC SECTION." & vbLf
    s = s & "       ERRHND-000." & vbLf
    s = s & "           MOVE W-W TO W-OUT." & vbLf
    s = s & "           PERFORM U-ABEND-PROC." & vbLf
    s = s & "       ERRHND-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       TAIL-SEC SECTION." & vbLf
    s = s & "       TAIL-000." & vbLf
    s = s & "           IF W-W = 'A'" & vbLf
    s = s & "           THEN" & vbLf
    s = s & "               MOVE 'AA' TO W-OUT" & vbLf
    s = s & "           ELSE" & vbLf
    s = s & "               MOVE 'BB' TO W-OUT" & vbLf
    s = s & "           END-IF." & vbLf
    s = s & "       TAIL-999." & vbLf
    s = s & "           EXIT." & vbLf
    s = s & "       U-ABEND-PROC SECTION." & vbLf
    s = s & "       UAB-000." & vbLf
    s = s & "           CALL 'ERRSUB'." & vbLf
    s = s & "       UAB-999." & vbLf
    s = s & "           EXIT." & vbLf

    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(s, New Collection)
    TestRunner.Assert_Equal CLng(0), CLng(UncovCount_(flow)), "transitive-abend: every arm covered"
    Dim tok As String
    tok = ArmTokenByDisp_(flow, "IF W-W = 'A' [THEN]")
    TestRunner.Assert_True Len(tok) > 0, "found the downstream TAIL arm"
    TestRunner.Assert_True NormCovers_(flow, tok), "downstream arm reachable (handler no longer blocks it)"
End Sub
