Attribute VB_Name = "CobolFlow"
' CobolFlow - ver3.0 P2: execution-flow path enumeration and test-case
' (scenario) generation. VBA port of the validated PS oracle.
'
' Model (spec 2026-06-10):
'   - Paths start at the procedure entry section, follow PERFORM call/return,
'     and stop at terminators: GOBACK / STOP RUN / EXIT PROGRAM actions, or a
'     PERFORM of a registered terminator section (ABEND handler).
'   - Literal constant propagation prunes infeasible paths: items assigned a
'     literal are tracked per path (CALL USING params and computed targets are
'     invalidated); equality-style branch conditions on a known item keep only
'     the consistent arm. Pruning happens only when certain.
'   - Cases = C1 greedy minimal cover of branch arms over normal paths,
'     + one case per code-derived abend arm, + one synthesized "call target
'     fails" stub per external CALL (the registered terminator sections'
'     own CALLs are never reached, so they are excluded naturally).
'
' Analyze_Flow returns OrderedDict:
'   cases   - Collection of case OrderedDicts (see BuildCase_)
'   arms    - Collection of {Token, Line, Disp} for every branch arm
'   normalPaths / abendPaths - Long counts of surviving candidate walks
'   truncated - Boolean (always False since ver3.1: directed walks never fork)
'   entryName - String

Option Explicit

Private Const MAX_TRACES As Long = 200
Private mGotoSeen As OrderedDict   ' GO TO targets followed in this walk (once each)
Private mOps As Long   ' heartbeat counter: keep Excel responsive on big programs
' every cons node ever allocated this run - unlinked iteratively at the end,
' because letting VB tear down a long Prev chain recursively can blow the
' native stack (classic linked-list teardown crash, untrappable in VBA)
Private mConsReg As Collection
' PerformInto_ runs per trace per PERFORM: cache the node list per range and
' index owners by name, or big programs rescan the whole node table millions
' of times (the stage-3 hot spot on 1000+ line sources)
Private mRangeCache As OrderedDict
Private mOwnerIdx As OrderedDict

' ---- directed path construction (ver3.2) ----
' Enumerating all paths is exponential in the branch count; for C1 coverage
' we construct candidates directly: two seed walks (THEN/ELSE preference,
' ABEND-avoiding), then SWEEP rounds that each maximize uncovered-arm
' pickup (direct pick at the branch, ancestor weights elsewhere) keeping
' the case count near the optimal cover, then a targeted fallback walk per
' still-uncovered arm (abends / value-steered flag arms). Work is linear
' in arms x path length, validated by the ver4 PS oracle.
Private mNeed As OrderedDict       ' required arm tokens for the current walk
Private mPrefElse As Boolean       ' seed preference for unsteered branches
Private mMissed As Boolean         ' walk could not honor a required arm
Private mRecordEvents As Boolean   ' pass 2 records events; pass 1 arms only
Private mStopAtCall As String      ' pass 2 synth: end the walk at this CALL
Private mArmCtx As OrderedDict     ' arm token -> Collection of ancestor tokens
Private mCtxVisited As OrderedDict ' sections already context-walked
Private mCurTok As String          ' walk identity (for synth re-walk specs)
' value-driven steering: when an arm tests ITEM against a literal and the
' first walk is blocked by a propagated constant, retry once steering the
' branch chain of a literal MOVE that establishes a satisfying value
' (covers the common COBOL flag idiom: a sibling branch sets the flag a
' later IF/EVALUATE tests).
Private mAssignCtx As OrderedDict  ' "ITEM=VAL" -> reach ctx of a MOVE site
Private mUnsetCtx As OrderedDict   ' ITEM -> reach ctx of a site making it unknown
Private mArmSteer As OrderedDict   ' arm token -> {Item, Val, Mode eq/ne}
' uncovered-arm diagnostics (reason codes; the matrix renders them in JP):
'   "noctx"           - arm has no reach context (never seen by the ctx walk)
'   "conflict|<cond>|<tried/nosite/nosteer>" - required arm infeasible
'                       under propagated constants (+ steering status)
'   "dead"            - walk died before reaching a terminator
Private mArmDiag As OrderedDict
Private mMissCond As String        ' condition that blocked the current walk
Private mMissTok As String         ' the needed arm that conflicted
Private mBlockTok As String        ' blocker of the last WalkTo_ first attempt
Private mBlockCond As String       ' its condition (attempt-1 snapshot - the
                                   ' retries overwrite mMissCond/mMissTok)
' havoc retry: treat this item as unknown for one walk. Used when a tested
' flag has a NON-LITERAL setter (value domain not statically known) but the
' single loop pass runs the setter AFTER the test (read-ahead idiom), so
' neither value- nor unset-steering can reorder the walk past the conflict.
Private mHavocItem As String

Private mNodes As Collection      ' AST root nodes
Private mOwners As Collection     ' {name,line,kind,ownerEnd,secEnd} sorted
Private mCut As OrderedDict       ' plain-PERFORM target names
Private mGotoCut As OrderedDict    ' GO TO target names (orphan-root exclusion only; never caps walk ranges)
Private mExecClears As OrderedDict ' EXEC line -> fields cleared by the literal MOVE right before it (DB-result pre-clear)
Private mForce As Boolean          ' fallback: take a value-conflicted arm anyway (data-driven blocker - exists therefore testable)
Private mForcedArm As OrderedDict   ' arm token -> the blocker condition that was force-satisfied (precondition mark)
Private mTermSecs As OrderedDict  ' registered terminator section names
Private mDesc As OrderedDict      ' item -> Collection of descendant names
Private mAnc As OrderedDict       ' item -> Collection of ancestor names
Private mCondItems As OrderedDict ' identifiers used in any branch condition
Private mSynth As OrderedDict     ' call target -> {Line, Target, ArmsAt}
Private mCurEntrySec As String     ' entry name the current walk starts from (stamped on candidates for pass-2 replay)
Private mEntryFallsThru As Boolean ' main entry has no GOBACK/STOP RUN/EXIT PROGRAM: reaching its end = implicit normal return
                                  ' (first site + the arm prefix at that point;
                                  '  pass 2 replays the prefix and stops at the call)
Private mTruncated As Boolean
' minimal-case sweep: each sweep walk grabs as many uncovered arms as it
' can (direct pick at the branch, ancestor weights toward uncovered arms
' elsewhere), so the candidate set stays near the optimal cover size.
Private mSweep As Boolean
Private mSweepUncov As OrderedDict ' arm tokens still uncovered this round
Private mSweepW As OrderedDict     ' ancestor-token weights toward uncovered
' pass-2 replay: reproduce a selected pass-1 walk EXACTLY by consuming its
' recorded arm sequence at each branch (deterministic regardless of sweep
' or steering state; also reproduces synthesized stop-at-call prefixes).
Private mReplayList As Collection
Private mReplayIdx As Long

'======================================================================
' Public entry
'======================================================================
Public Function Analyze_Flow(ByVal src As String, ByVal termSections As Collection) As OrderedDict
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Set mNodes = CobolParser.Get_CobolNodes(lines)
    Dim struct As OrderedDict
    Set struct = CobolParser.Get_ProgramStructure(lines)

    Set mTermSecs = New OrderedDict
    Dim v As Variant
    If Not termSections Is Nothing Then
        For Each v In termSections
            If Len(Trim$(CStr(v))) > 0 Then
                If Not mTermSecs.Exists(UCase$(Trim$(CStr(v)))) Then mTermSecs.Add UCase$(Trim$(CStr(v))), True
            End If
        Next v
    End If

    BuildOwners_ struct

    ' auto-detect terminator sections by naming convention (name contains
    ' "ABEND") so the common case needs no manual registration; the control
    ' sheet cells (B24:B29) remain for other naming schemes. What was applied
    ' is reported via "termsApplied" and shown on the case sheet.
    Dim termsApplied As Collection
    Set termsApplied = New Collection
    Dim tk As Collection
    Set tk = mTermSecs.Keys
    For Each v In tk
        termsApplied.Add MakeTermInfo_(CStr(v), "manual")
    Next v
    Set mOwnerIdx = New OrderedDict
    Set mRangeCache = New OrderedDict
    Dim ow As OrderedDict
    For Each ow In mOwners
        If Not mOwnerIdx.Exists(CStr(ow.Item("name"))) Then mOwnerIdx.Add CStr(ow.Item("name")), ow
        If CStr(ow.Item("kind")) = "section" And InStr(CStr(ow.Item("name")), "ABEND") > 0 Then
            If Not mTermSecs.Exists(CStr(ow.Item("name"))) Then
                mTermSecs.Add CStr(ow.Item("name")), True
                termsApplied.Add MakeTermInfo_(CStr(ow.Item("name")), "auto")
            End If
        End If
    Next ow

    BuildCut_
    Set mExecClears = New OrderedDict
    BuildExecClears_ mNodes

    ' transitive terminators: a SECTION that unconditionally reaches a
    ' terminator (a top-level PERFORM of a terminator, or GOBACK, before any
    ' top-level branch) is itself a terminator - e.g. an error handler that
    ' always PERFORMs the abend section. Without this, abend-avoidance walks
    ' into such a PERFORM, dies, and blocks every downstream arm. Fixpoint.
    ' MUST run after BuildCut_: SectionAlwaysTerminates_ -> RangeNodes_ ->
    ' CapHi_ reads mCut, which BuildCut_ builds (else error 91 on Nothing).
    Dim tChanged As Boolean, ow2 As OrderedDict, nmT As String
    Do
        tChanged = False
        For Each ow2 In mOwners
            If CStr(ow2.Item("kind")) = "section" Then
                nmT = CStr(ow2.Item("name"))
                If Not mTermSecs.Exists(nmT) Then
                    If SectionAlwaysTerminates_(CLng(ow2.Item("line")), CLng(ow2.Item("secEnd"))) Then
                        mTermSecs.Add nmT, True
                        termsApplied.Add MakeTermInfo_(nmT, "auto")
                        tChanged = True
                    End If
                End If
            End If
        Next ow2
    Loop While tChanged
    BuildGroupMaps_ lines
    BuildCondItems_ mNodes
    Set mSynth = New OrderedDict
    mTruncated = False
    mOps = 0
    Set mConsReg = New Collection

    ' stage labels: comment with full-width parens just above a section header
    Dim secLabels As OrderedDict
    Set secLabels = BuildSecLabels_(norm)

    Dim arms As Collection
    Set arms = New Collection
    CollectArms_ mNodes, arms

    ' pass 1: candidate walks, arms/term only (no events, no forking).
    ' two seeds, then sweep rounds that each grab as many uncovered arms
    ' as possible (keeps the case count near the optimal cover), then a
    ' targeted fallback walk per still-uncovered arm (abends / steering).
    Dim entry As OrderedDict
    Set entry = FindEntry_()
    ' scan the SAME range the top-level walk uses (CapHi_-capped, not the raw
    ' secEnd) so a GOBACK in a performed paragraph past the cap does not hide
    ' the fall-through of the actual main flow
    mEntryFallsThru = False
    If Not entry Is Nothing Then
        mEntryFallsThru = Not LinesHaveNormalTerm_(lines, CLng(entry.Item("line")), _
                              CapHi_(CLng(entry.Item("line")), CLng(entry.Item("secEnd"))))
    End If
    Dim cands As Collection
    Set cands = New Collection
    Dim covered As OrderedDict
    Set covered = New OrderedDict
    Dim w As OrderedDict
    If Not entry Is Nothing And mNodes.Count > 0 Then
        BuildArmCtx_ entry, secLabels
        mCurEntrySec = CStr(entry.Item("name"))
        mRecordEvents = False
        mStopAtCall = ""
        Set mReplayList = Nothing
        Set w = WalkTo_("", False, entry, secLabels)
        AddCand_ cands, covered, w
        Set w = WalkTo_("", True, entry, secLabels)
        AddCand_ cands, covered, w

        Dim rounds As Long
        rounds = 0
        Do
            If Not BuildSweepState_(arms, covered) Then Exit Do
            Set w = SweepWalk_(entry, secLabels)
            If CStr(w.Item("Term")) = "" Then Exit Do
            If SweepGain_(w, covered) = 0 Then Exit Do
            AddCand_ cands, covered, w
            rounds = rounds + 1
            If rounds > arms.Count Then Exit Do
        Loop

        ' call edges from the ver2 extractor: lets the noctx diagnostics
        ' name the unreached section AND who the call graph says invokes
        ' it - distinguishing a flow-walker gap from genuinely dead code
        Dim callG As OrderedDict
        On Error Resume Next
        Set callG = CobolParser.Get_CallRelationships(lines, struct)
        On Error GoTo 0

        Dim a As OrderedDict, tok As String
        Set mArmDiag = New OrderedDict
        Set mForcedArm = New OrderedDict
        For Each a In arms
            tok = CStr(a.Item("Token"))
            If Not covered.Exists(tok) Then
                If Not mArmCtx.Exists(tok) Then
                    mArmDiag.Add tok, "noctx|" & NoCtxDetail_(CLng(a.Item("Line")), callG, lines)
                Else
                    Set w = WalkTo_(tok, False, entry, secLabels)
                    If CStr(w.Item("Term")) <> "" And Not CBool(w.Item("Missed")) Then
                        AddCand_ cands, covered, w
                    ElseIf CBool(w.Item("Missed")) And Len(mBlockCond) > 0 Then
                        ' exists therefore testable: the blocker is an
                        ' assignment-pinned value (often a DB/CALL result or a
                        ' data-driven intermediate) the engine cannot satisfy via
                        ' steering. Force the arm so the branch is covered, and
                        ' remember the blocker condition as the case precondition
                        ' (driver/DB must set it up).
                        Dim blkCond As String, blkTok As String
                        blkCond = mBlockCond
                        blkTok = mBlockTok
                        mForce = True
                        Dim wFf As OrderedDict
                        Set wFf = WalkTo_(tok, False, entry, secLabels)
                        mForce = False
                        If CStr(wFf.Item("Term")) <> "" And Not CBool(wFf.Item("Missed")) Then
                            AddCand_ cands, covered, wFf
                            If Not mForcedArm.Exists(tok) Then mForcedArm.Add tok, blkCond
                        Else
                            ' force failed too: annotate what steering had to work
                            ' with (suffix considers the target arm AND the blocker)
                            Dim sfx As String, sk As Variant
                            sfx = "nosteer"
                            For Each sk In Array(tok, blkTok)
                                If Len(CStr(sk)) > 0 Then
                                    If mArmSteer.Exists(CStr(sk)) Then
                                        If Not SteerChain_(CStr(sk)) Is Nothing Then
                                            sfx = "tried"
                                            Exit For
                                        ElseIf sfx = "nosteer" Then
                                            sfx = "nosite"
                                        End If
                                    End If
                                End If
                            Next sk
                            mArmDiag.Add tok, "conflict|" & blkCond & "|" & sfx
                        End If
                    Else
                        mArmDiag.Add tok, "dead"
                    End If
                End If
            End If
        Next a

        ' second pass: an arm can be "covered" only because a seed reached
        ' it then ran into a DEEP abend (one beyond the shallow abend-
        ' avoidance), so it has no NORMAL case - it shows up only under an
        ' abnormal column. For every arm covered abnormally but not by a normal
        ' (goback) candidate, try a targeted normal walk; add it if it ends
        ' normally. Genuinely abend-only arms fail here (their walk abends)
        ' and stay abnormal-only.
        Dim normCov As OrderedDict, ca As OrderedDict, cv As Variant
        Dim a2 As OrderedDict, wN As OrderedDict, t2 As String
        Set normCov = New OrderedDict
        For Each ca In cands
            If CStr(ca.Item("Term")) = "goback" Then
                For Each cv In ca.Item("ArmsL")
                    If Not normCov.Exists(CStr(cv)) Then normCov.Add CStr(cv), True
                Next cv
            End If
        Next ca
        For Each a2 In arms
            t2 = CStr(a2.Item("Token"))
            If mArmCtx.Exists(t2) And covered.Exists(t2) And Not normCov.Exists(t2) Then
                Set wN = WalkTo_(t2, False, entry, secLabels)
                If CStr(wN.Item("Term")) = "goback" And Not CBool(wN.Item("Missed")) Then
                    AddCand_ cands, covered, wN
                    For Each cv In wN.Item("ArmsL")
                        If Not normCov.Exists(CStr(cv)) Then normCov.Add CStr(cv), True
                    Next cv
                End If
            End If
        Next a2

        ' ver3.9 unit-test entries: a procedure SECTION the main flow never
        ' PERFORMs (an "orphan" - e.g. an update-via-CALL path superseded by an
        ' inline DB verb) is still unit-testable: a driver PERFORMs it directly
        ' and stubs its CALLs, so its branches DO get cases. Walk each such
        ' section as its own entry. Synth (CALL-failure) specs are frozen across
        ' this phase, so orphan cases are normal + code-abend only, and each
        ' carries its entry section for the pass-2 replay.
        Dim oSec As OrderedDict, synthSnap As OrderedDict, r2 As Long
        Dim ao As OrderedDict, tko As String
        Set synthSnap = mSynth
        For Each oSec In OrphanRoots_(entry)
            If SecHasUncoveredArm_(oSec, arms, covered) Then
                Set mSynth = New OrderedDict
                mCurEntrySec = CStr(oSec.Item("name"))
                BuildArmCtx_ oSec, secLabels
                AddUnitCand_ cands, covered, WalkTo_("", False, oSec, secLabels)
                AddUnitCand_ cands, covered, WalkTo_("", True, oSec, secLabels)
                r2 = 0
                Do
                    If Not BuildSweepState_(arms, covered) Then Exit Do
                    Set w = SweepWalk_(oSec, secLabels)
                    If SweepGain_(w, covered) = 0 Then Exit Do
                    AddUnitCand_ cands, covered, w
                    r2 = r2 + 1
                    If r2 > arms.Count Then Exit Do
                Loop
                For Each ao In arms
                    tko = CStr(ao.Item("Token"))
                    If Not covered.Exists(tko) And mArmCtx.Exists(tko) Then
                        Set w = WalkTo_(tko, False, oSec, secLabels)
                        If Not CBool(w.Item("Missed")) Then AddUnitCand_ cands, covered, w
                    End If
                Next ao
            End If
        Next oSec
        Set mSynth = synthSnap
        mCurEntrySec = CStr(entry.Item("name"))
    End If
    If mArmDiag Is Nothing Then Set mArmDiag = New OrderedDict

    ' split normal / abend candidates
    Dim normals As Collection, abends As Collection, tr As OrderedDict
    Set normals = New Collection
    Set abends = New Collection
    For Each tr In cands
        If CStr(tr.Item("Term")) = "goback" Then
            normals.Add tr
        ElseIf Left$(CStr(tr.Item("Term")), 6) = "abend:" Then
            abends.Add tr
        End If
    Next tr

    Dim specs As Collection
    Set specs = SelectCases_(normals, abends, arms)

    ' pass 2: replay each selected case's exact arm sequence with events
    Dim cases As Collection
    Set cases = New Collection
    Dim sp As OrderedDict, w2 As OrderedDict
    mRecordEvents = True
    For Each sp In specs
        If CStr(sp.Item("Kind")) = "synth" Then
            mStopAtCall = CStr(sp.Item("SynthTarget"))
        Else
            mStopAtCall = ""
        End If
        Dim epSec As OrderedDict, eo As OrderedDict, spUnit As Boolean
        Set epSec = entry
        spUnit = False
        If sp.Exists("EntrySec") Then
            If CStr(sp.Item("EntrySec")) <> CStr(entry.Item("name")) Then
                Set eo = OwnerByName_(CStr(sp.Item("EntrySec")))
                If Not eo Is Nothing Then
                    Set epSec = eo
                    spUnit = True
                End If
            End If
        End If
        Set w2 = ReplayWalk_(sp.Item("NeedList"), epSec, secLabels)
        If spUnit And CStr(w2.Item("Term")) = "" Then w2.Add "Term", "goback"
        cases.Add BuildCase_(w2, CStr(sp.Item("Id")), CStr(sp.Item("Kind")), CLng(sp.Item("KindSerial")))
    Next sp
    mStopAtCall = ""
    mRecordEvents = False

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "cases", cases
    result.Add "arms", arms
    ' SECTION ranges for sheet renderers (owning-section lookup by line);
    ' data-division sections are excluded, same as the entry detection.
    ' "note" = the comment above the section header (kanji description).
    Dim cmtMap As OrderedDict, cmtE As OrderedDict, codeSet As OrderedDict, lnE As OrderedDict
    Set cmtMap = New OrderedDict
    If norm.Exists("Comments") Then
        For Each cmtE In norm.Item("Comments")
            If Not cmtMap.Exists(CStr(cmtE.Item("Number"))) Then _
                cmtMap.Add CStr(cmtE.Item("Number")), CStr(cmtE.Item("Text"))
        Next cmtE
    End If
    ' code-line set so SecNote_ can skip blank gaps yet stop at real code
    Set codeSet = New OrderedDict
    For Each lnE In lines
        If Not codeSet.Exists(CStr(lnE.Item("Number"))) Then codeSet.Add CStr(lnE.Item("Number")), True
    Next lnE
    Dim secList As Collection, sd As OrderedDict
    Set secList = New Collection
    For Each ow In mOwners
        If CStr(ow.Item("kind")) = "section" Then
            If Not IsDataSection_(CStr(ow.Item("name"))) Then
                Set sd = New OrderedDict
                sd.Add "name", CStr(ow.Item("name"))
                sd.Add "line", CLng(ow.Item("line"))
                sd.Add "secEnd", CLng(ow.Item("secEnd"))
                sd.Add "note", SecNote_(cmtMap, codeSet, CLng(ow.Item("line")))
                secList.Add sd
            End If
        End If
    Next ow
    result.Add "sections", secList
    result.Add "normalPaths", normals.Count
    result.Add "abendPaths", abends.Count
    result.Add "truncated", mTruncated
    result.Add "descMap", mDesc   ' item -> descendants (for downstream IO derivation)
    result.Add "termsApplied", termsApplied
    result.Add "armDiag", mArmDiag   ' uncovered-arm reason codes (matrix)
    If mForcedArm Is Nothing Then Set mForcedArm = New OrderedDict
    result.Add "forcedArms", mForcedArm   ' force-covered arm -> blocker condition (driver/DB must set up)
    If entry Is Nothing Then
        result.Add "entryName", ""
    Else
        result.Add "entryName", CStr(entry.Item("name"))
    End If
    UnlinkCons_   ' cases hold materialized Collections by now
    Set mOwnerIdx = Nothing    ' per-run caches must not outlive the run
    Set mRangeCache = Nothing
    Set Analyze_Flow = result
End Function

Private Function MakeTermInfo_(ByVal nm As String, ByVal src As String) As OrderedDict
    Dim t As OrderedDict
    Set t = New OrderedDict
    t.Add "name", nm
    t.Add "source", src   ' "auto" (ABEND naming) / "manual" (control sheet)
    Set MakeTermInfo_ = t
End Function

'======================================================================
' Directed path construction
'======================================================================
' One walk = one trace: branches on the way to the target arm are steered
' (mNeed), all others take a default feasible arm. Pass 2 reproduces a
' selected walk by replaying its recorded arm sequence (ReplayWalk_), so
' the result only needs its Arms/Term - no walk-identity tags.
Private Function WalkTo_(ByVal token As String, ByVal prefElse As Boolean, _
                         ByVal entry As OrderedDict, ByVal secLabels As OrderedDict) As OrderedDict
    Dim res As OrderedDict
    Set res = TryWalk_(token, prefElse, entry, secLabels, Nothing, "")
    mBlockTok = mMissTok
    mBlockCond = mMissCond
    If Not CBool(res.Item("Missed")) Or Len(token) = 0 Then
        Set WalkTo_ = res
        Exit Function
    End If
    ' retries are keyed first on the TARGET arm's steer info, then on the
    ' BLOCKING arm's: the target often has none (compound condition) while
    ' the ancestor that actually conflicted is steerable / havoc-able.
    ' Per key: (a) value/unset chain retry - steer toward a site that
    ' satisfies or unsets the tested item; (b) havoc retry - walk treating
    ' the item as unknown, gated on a non-literal setter existing.
    Dim keys As Collection, rk As Variant
    Set keys = New Collection
    keys.Add token
    If Len(mBlockTok) > 0 And mBlockTok <> token Then keys.Add mBlockTok
    For Each rk In keys
        If mArmSteer.Exists(CStr(rk)) Then
            Dim extra As Collection
            Set extra = SteerChain_(CStr(rk))
            If Not extra Is Nothing Then
                Dim res2 As OrderedDict
                Set res2 = TryWalk_(token, prefElse, entry, secLabels, extra, "")
                If Not CBool(res2.Item("Missed")) Then
                    Set WalkTo_ = res2
                    Exit Function
                End If
            End If
            Dim si3 As OrderedDict
            Set si3 = mArmSteer.Item(CStr(rk))
            If mUnsetCtx.Exists(CStr(si3.Item("Item"))) Then
                Dim res3 As OrderedDict
                Set res3 = TryWalk_(token, prefElse, entry, secLabels, Nothing, CStr(si3.Item("Item")))
                If Not CBool(res3.Item("Missed")) Then
                    Set WalkTo_ = res3
                    Exit Function
                End If
            End If
        End If
    Next rk
    Set WalkTo_ = res
End Function

Private Function TryWalk_(ByVal token As String, ByVal prefElse As Boolean, _
                          ByVal entry As OrderedDict, ByVal secLabels As OrderedDict, _
                          ByVal extraNeed As Collection, ByVal havocItem As String) As OrderedDict
    Set mNeed = New OrderedDict
    Dim v As Variant
    If Len(token) > 0 Then
        For Each v In mArmCtx.Item(token)
            mNeed.Add CStr(v), True
        Next v
        mNeed.Add token, True
    End If
    If Not extraNeed Is Nothing Then
        For Each v In extraNeed
            mNeed.Add CStr(v), True
        Next v
    End If
    mPrefElse = prefElse
    mMissed = False
    mMissCond = ""
    mMissTok = ""
    mCurTok = token
    mSweep = False
    mHavocItem = havocItem
    Set mGotoSeen = New OrderedDict
    Set mReplayList = Nothing

    Dim res As OrderedDict
    Set res = RunWalk_(entry, secLabels)
    mHavocItem = ""
    If Len(token) > 0 And Not mMissed Then
        If Not ArmOnTrace_(res, token) Then mMissed = True
    End If
    res.Add "Missed", mMissed
    Set TryWalk_ = res
End Function

' sweep walk: no steering, choices driven by mSweepUncov/mSweepW
Private Function SweepWalk_(ByVal entry As OrderedDict, ByVal secLabels As OrderedDict) As OrderedDict
    Set mNeed = New OrderedDict
    mPrefElse = False
    mMissed = False
    mCurTok = ""
    Set mGotoSeen = New OrderedDict
    Set mReplayList = Nothing
    mSweep = True
    Dim res As OrderedDict
    Set res = RunWalk_(entry, secLabels)
    mSweep = False
    res.Add "Missed", False
    Set SweepWalk_ = res
End Function

' pass-2 walk: consume the recorded arm sequence at each branch
Private Function ReplayWalk_(ByVal armSeq As Collection, ByVal entry As OrderedDict, _
                             ByVal secLabels As OrderedDict) As OrderedDict
    Set mNeed = New OrderedDict
    mPrefElse = False
    mMissed = False
    mCurTok = ""
    mSweep = False
    Set mGotoSeen = New OrderedDict
    Set mReplayList = armSeq
    mReplayIdx = 1
    Set ReplayWalk_ = RunWalk_(entry, secLabels)
    Set mReplayList = Nothing
End Function

' shared walk core: one trace through the inline expansion from the entry
Private Function RunWalk_(ByVal entry As OrderedDict, ByVal secLabels As OrderedDict) As OrderedDict
    Dim t0 As OrderedDict
    Set t0 = NewTrace_()
    AddEnterEvent_ t0, CStr(entry.Item("name")), secLabels, CLng(entry.Item("line"))
    Dim seed As Collection, stack As Collection
    Set seed = New Collection
    seed.Add t0
    Set stack = New Collection
    stack.Add CStr(entry.Item("name"))
    Dim outs As Collection
    Set outs = ApplyRange_(CLng(entry.Item("line")), CLng(entry.Item("secEnd")), stack, seed, secLabels)
    If outs.Count > 0 Then
        Set RunWalk_ = outs(1)
        ' main entry with no explicit terminator: reaching the end of the
        ' top-level range is an implicit normal return (control falls off the
        ' procedure division). Promote so normal cases generate and the sweep
        ' keeps flowing instead of treating every fall-through as a dead walk.
        If mEntryFallsThru And CStr(RunWalk_.Item("Term")) = "" Then RunWalk_.Add "Term", "goback"
    Else
        Set RunWalk_ = t0   ' walk died (no feasible arm) - Term "", dropped
    End If
End Function

' the extra need-chain for a value-driven retry, or Nothing when the arm
' has no literal steering info / no satisfying assignment site exists
Private Function SteerChain_(ByVal token As String) As Collection
    Set SteerChain_ = Nothing
    If Not mArmSteer.Exists(token) Then Exit Function
    Dim si As OrderedDict
    Set si = mArmSteer.Item(token)
    Dim key As String, vals() As String, i As Long
    key = ""
    vals = Split(CStr(si.Item("Val")), "|")
    If CStr(si.Item("Mode")) = "eq" Then
        ' establish ANY of the listed values
        For i = LBound(vals) To UBound(vals)
            If mAssignCtx.Exists(CStr(si.Item("Item")) & "=" & vals(i)) Then
                key = CStr(si.Item("Item")) & "=" & vals(i)
                Exit For
            End If
        Next i
    Else
        ' any literal assignment of a value OUTSIDE the list
        Dim ks As Collection, v As Variant, pre As String, hitList As Boolean
        pre = CStr(si.Item("Item")) & "="
        Set ks = mAssignCtx.Keys
        For Each v In ks
            If Left$(CStr(v), Len(pre)) = pre Then
                hitList = False
                For i = LBound(vals) To UBound(vals)
                    If CStr(v) = pre & vals(i) Then hitList = True
                Next i
                If Not hitList Then
                    key = CStr(v)
                    Exit For
                End If
            End If
        Next v
    End If
    If Len(key) > 0 Then
        Set SteerChain_ = mAssignCtx.Item(key)
        Exit Function
    End If
    ' unset-steering fallback: no literal site matches, but a site that
    ' makes the item UNKNOWN unblocks the constant-propagation conflict
    If mUnsetCtx.Exists(CStr(si.Item("Item"))) Then
        Set SteerChain_ = mUnsetCtx.Item(CStr(si.Item("Item")))
    End If
End Function

' literal-equality steering info for an IF node's arms. "Val" carries a
' pipe-joined value LIST: eq = establish any listed value, ne = establish
' a value outside the list. OR-lists of equalities on the same item are
' the exact shape TestFeasible_ prunes, so they need steering too.
Private Sub SteerInfo_(ByVal cond As String, ByVal nodeId As String)
    Dim c0 As String
    c0 = Trim$(cond)
    If InStr(c0, " AND ") > 0 Then Exit Sub
    Static rxEq As Object, rxNe As Object
    If rxEq Is Nothing Then
        Set rxEq = CreateObject("VBScript.RegExp"): rxEq.Pattern = "^([A-Z0-9-]+)\s*=\s*(.+)$": rxEq.IgnoreCase = False
        Set rxNe = CreateObject("VBScript.RegExp"): rxNe.Pattern = "^([A-Z0-9-]+)\s+NOT\s*=\s*(.+)$": rxNe.IgnoreCase = False
    End If
    Dim m As Object, lit As String
    If InStr(c0, " OR ") > 0 Then
        Dim parts() As String, i As Long, item As String, vals As String
        parts = Split(c0, " OR ")
        item = ""
        vals = ""
        For i = LBound(parts) To UBound(parts)
            Set m = rxEq.Execute(Trim$(parts(i)))
            If m.Count = 0 Then Exit Sub
            If Not GetLiteral_(m.Item(0).SubMatches(1), lit) Then Exit Sub
            If i = LBound(parts) Then
                item = m.Item(0).SubMatches(0)
            ElseIf item <> m.Item(0).SubMatches(0) Then
                Exit Sub
            End If
            If Len(vals) > 0 Then vals = vals & "|"
            vals = vals & NormVal_(lit)
        Next i
        AddSteer_ nodeId & ":then", item, vals, "eq"
        AddSteer_ nodeId & ":else", item, vals, "ne"
        Exit Sub
    End If
    Set m = rxNe.Execute(c0)
    If m.Count > 0 Then
        If GetLiteral_(m.Item(0).SubMatches(1), lit) Then
            AddSteer_ nodeId & ":then", m.Item(0).SubMatches(0), NormVal_(lit), "ne"
            AddSteer_ nodeId & ":else", m.Item(0).SubMatches(0), NormVal_(lit), "eq"
        End If
        Exit Sub
    End If
    Set m = rxEq.Execute(c0)
    If m.Count > 0 Then
        If GetLiteral_(m.Item(0).SubMatches(1), lit) Then
            AddSteer_ nodeId & ":then", m.Item(0).SubMatches(0), NormVal_(lit), "eq"
            AddSteer_ nodeId & ":else", m.Item(0).SubMatches(0), NormVal_(lit), "ne"
        End If
    End If
End Sub

Private Sub AddSteer_(ByVal token As String, ByVal item As String, ByVal val As String, ByVal mode As String)
    If mArmSteer.Exists(token) Then Exit Sub
    Dim s As OrderedDict
    Set s = New OrderedDict
    s.Add "Item", item
    s.Add "Val", val
    s.Add "Mode", mode
    mArmSteer.Add token, s
End Sub

' record assignment sites with their reach context (first site wins):
' literal MOVEs feed value steering; non-literal MOVE / INITIALIZE make the
' item UNKNOWN and feed the unset-steering fallback (the common DB-flag
' idiom: MOVE <record-field> TO F-XXX in a sibling branch).
Private Sub RegisterAssign_(ByVal lbl As String, ByVal ctx As Collection)
    Static rxML As Object, rxMA As Object
    If rxML Is Nothing Then
        Set rxML = CreateObject("VBScript.RegExp")
        rxML.Pattern = "^MOVE\s+('[^']*'|[0-9]+|ZEROS?|ZEROES|SPACES?)\s+TO\s+([A-Z0-9-]+(\s+[A-Z0-9-]+)*)$"
        rxML.IgnoreCase = False
        Set rxMA = CreateObject("VBScript.RegExp")
        rxMA.Pattern = "^MOVE\s+(.+?)\s+TO\s+([A-Z0-9-]+(\s+[A-Z0-9-]+)*)$"
        rxMA.IgnoreCase = False
    End If
    Dim m As Object, dts() As String, i As Long, k As String
    ' CALL may modify its USING args (ADABAS RC etc.) - unset sites so
    ' steering/havoc can route a blocked test through the call
    If Left$(lbl, 5) = "CALL " Then
        Static rxCallU As Object
        If rxCallU Is Nothing Then
            Set rxCallU = CreateObject("VBScript.RegExp")
            rxCallU.Pattern = "^CALL\s+'([A-Z0-9-]+)'(\s+USING\s+(.+))?$"
            rxCallU.IgnoreCase = False
        End If
        Set m = rxCallU.Execute(lbl)
        If m.Count = 0 Then
            Static rxCallUV As Object
            If rxCallUV Is Nothing Then
                Set rxCallUV = CreateObject("VBScript.RegExp")
                rxCallUV.Pattern = "^CALL\s+([A-Z0-9-]+)(\s+USING\s+(.+))?$"
                rxCallUV.IgnoreCase = False
            End If
            Set m = rxCallUV.Execute(lbl)
        End If
        If m.Count > 0 Then
            Dim cps As String
            cps = m.Item(0).SubMatches(2)
            If Len(cps) > 0 Then
                dts = Split(Trim$(cps), " ")
                For i = LBound(dts) To UBound(dts)
                    k = UCase$(Trim$(dts(i)))
                    If Len(k) > 0 Then
                        If Not mUnsetCtx.Exists(k) Then mUnsetCtx.Add k, CopyCtx_(ctx)
                        ' the RC idiom passes a GROUP and tests a field
                        ' inside it - register subordinates too
                        If mDesc.Exists(k) Then
                            Dim dv As Variant
                            For Each dv In mDesc.Item(k)
                                If Not mUnsetCtx.Exists(CStr(dv)) Then mUnsetCtx.Add CStr(dv), CopyCtx_(ctx)
                            Next dv
                        End If
                    End If
                Next i
            End If
        End If
        Exit Sub
    End If
    ' arithmetic targets become unknown - register as unset sites
    If Left$(lbl, 4) = "ADD " Or Left$(lbl, 9) = "SUBTRACT " Or _
       Left$(lbl, 9) = "MULTIPLY " Or Left$(lbl, 7) = "DIVIDE " Then
        Dim ac As Collection, avv As Variant
        Set ac = ArithTargets_(lbl)
        For Each avv In ac
            UnsetWithDesc_ CStr(avv), ctx
        Next avv
        Exit Sub
    End If
    ' READ <file> INTO <item>: the record lands in <item> - unknown value
    ' (bare READ: FD record area not tracked - accepted limitation)
    If Left$(lbl, 5) = "READ " Then
        i = InStr(lbl, " INTO ")
        If i > 0 Then
            k = Trim$(Mid$(lbl, i + 6))
            i = InStr(k, " ")
            If i > 0 Then k = Left$(k, i - 1)
            k = UCase$(k)
            If Len(k) > 0 Then UnsetWithDesc_ k, ctx
        End If
        Exit Sub
    End If
    If Left$(lbl, 11) = "INITIALIZE " Then
        k = Trim$(Mid$(lbl, 12))
        i = InStr(k, " ")
        If i > 0 Then k = Left$(k, i - 1)
        If Len(k) > 0 Then UnsetWithDesc_ k, ctx
        Exit Sub
    End If
    If Left$(lbl, 5) <> "MOVE " Then Exit Sub
    Set m = rxML.Execute(lbl)
    If m.Count > 0 Then
        Dim lit As String
        If Not GetLiteral_(m.Item(0).SubMatches(0), lit) Then Exit Sub
        dts = Split(Trim$(m.Item(0).SubMatches(1)), " ")
        For i = LBound(dts) To UBound(dts)
            If Len(Trim$(dts(i))) > 0 Then
                k = Trim$(dts(i)) & "=" & NormVal_(lit)
                If Not mAssignCtx.Exists(k) Then mAssignCtx.Add k, CopyCtx_(ctx)
            End If
        Next i
        Exit Sub
    End If
    Set m = rxMA.Execute(lbl)
    If m.Count > 0 Then
        dts = Split(Trim$(m.Item(0).SubMatches(1)), " ")
        For i = LBound(dts) To UBound(dts)
            If Len(Trim$(dts(i))) > 0 Then UnsetWithDesc_ Trim$(dts(i)), ctx
        Next i
    End If
End Sub

' an unknown-value write to nm unsets nm AND its subordinate fields
' (group MOVE / group CALL param / READ INTO a group record).
' NOTE: INITIALIZE targets also land here - strictly its values ARE known
' (SPACE/ZERO), so this over-approximates: a steered retry may route
' through an INITIALIZE that cannot produce the wanted value. Accepted -
' it only ever ADDS candidate routes, and the conservative walker
' invalidation treats INITIALIZE the same way.
Private Sub UnsetWithDesc_(ByVal nm As String, ByVal ctx As Collection)
    If Not mUnsetCtx.Exists(nm) Then mUnsetCtx.Add nm, CopyCtx_(ctx)
    If mDesc.Exists(nm) Then
        Dim dv As Variant
        For Each dv In mDesc.Item(nm)
            If Not mUnsetCtx.Exists(CStr(dv)) Then mUnsetCtx.Add CStr(dv), CopyCtx_(ctx)
        Next dv
    End If
End Sub

Private Function CopyCtx_(ByVal ctx As Collection) As Collection
    Dim c As Collection, v As Variant
    Set c = New Collection
    For Each v In ctx
        c.Add CStr(v)
    Next v
    Set CopyCtx_ = c
End Function

Private Function ArmOnTrace_(ByVal tr As OrderedDict, ByVal token As String) As Boolean
    Dim cur As ConsList
    Set cur = tr.Item("Arms")
    ArmOnTrace_ = False
    Do While Not cur Is Nothing
        If CStr(cur.V) = token Then
            ArmOnTrace_ = True
            Exit Function
        End If
        Set cur = cur.Prev
    Loop
End Function

' arm token -> chain of ancestor arm tokens required to reach it. One
' depth-first walk over the inline expansion; each section is context-walked
' once (the first reaching context wins), so the work stays linear.
Private Sub BuildArmCtx_(ByVal entry As OrderedDict, ByVal secLabels As OrderedDict)
    Set mArmCtx = New OrderedDict
    Set mCtxVisited = New OrderedDict
    Set mAssignCtx = New OrderedDict
    Set mUnsetCtx = New OrderedDict
    Set mArmSteer = New OrderedDict
    Dim stack As Collection
    Set stack = New Collection
    stack.Add CStr(entry.Item("name"))
    mCtxVisited.Add CStr(entry.Item("name")), True
    CtxWalk_ RangeNodes_(CLng(entry.Item("line")), CLng(entry.Item("secEnd"))), stack, New Collection
End Sub

Private Sub CtxWalk_(ByVal list As Collection, ByVal stack As Collection, ByVal ctx As Collection)
    Dim n As OrderedDict, t As String
    For Each n In list
        t = CStr(n.Item("type"))
        If t = "if" Then
            RegisterCtx_ CStr(n.Item("id")) & ":then", ctx
            RegisterCtx_ CStr(n.Item("id")) & ":else", ctx
            SteerInfo_ CStr(n.Item("condition")), CStr(n.Item("id"))
            CtxWalk_ n.Item("thenChildren"), stack, CtxPlus_(ctx, CStr(n.Item("id")) & ":then")
            CtxWalk_ n.Item("elseChildren"), stack, CtxPlus_(ctx, CStr(n.Item("id")) & ":else")
        ElseIf t = "evaluate" Then
            Dim cs As Collection, wi As Long, hasOther As Boolean, wlit As String
            Dim otherId As String, allLits As String
            Set cs = n.Item("cases")
            hasOther = False
            otherId = ""
            allLits = ""
            For wi = 1 To cs.Count
                RegisterCtx_ CStr(cs(wi).Item("id")), ctx
                If CStr(cs(wi).Item("condition")) = "OTHER" Then
                    hasOther = True
                    otherId = CStr(cs(wi).Item("id"))
                End If
                If GetLiteral_(CStr(cs(wi).Item("condition")), wlit) Then
                    AddSteer_ CStr(cs(wi).Item("id")), CStr(n.Item("expression")), NormVal_(wlit), "eq"
                    If Len(allLits) > 0 Then allLits = allLits & "|"
                    allLits = allLits & NormVal_(wlit)
                End If
                CtxWalk_ cs(wi).Item("children"), stack, CtxPlus_(ctx, CStr(cs(wi).Item("id")))
            Next wi
            ' OTHER / implicit skip need "a value outside every listed WHEN"
            If Len(allLits) > 0 Then
                If hasOther Then
                    AddSteer_ otherId, CStr(n.Item("expression")), allLits, "ne"
                Else
                    AddSteer_ CStr(n.Item("id")) & ":skip", CStr(n.Item("expression")), allLits, "ne"
                End If
            End If
            If Not hasOther Then RegisterCtx_ CStr(n.Item("id")) & ":skip", ctx
        ElseIf t = "search" Then
            If Not IsNull(n.Item("atEndLine")) Then
                RegisterCtx_ CStr(n.Item("id")) & ":atend", ctx
                CtxWalk_ n.Item("atEndChildren"), stack, CtxPlus_(ctx, CStr(n.Item("id")) & ":atend")
            Else
                RegisterCtx_ CStr(n.Item("id")) & ":skip", ctx
            End If
            Dim sc As Collection, si As Long
            Set sc = n.Item("cases")
            For si = 1 To sc.Count
                RegisterCtx_ CStr(sc(si).Item("id")), ctx
                CtxWalk_ sc(si).Item("children"), stack, CtxPlus_(ctx, CStr(sc(si).Item("id")))
            Next si
        ElseIf t = "action" Then
            CtxPerform_ CStr(n.Item("label")), stack, ctx
            CtxGoTo_ CStr(n.Item("label")), stack, ctx
            RegisterAssign_ CStr(n.Item("label")), ctx
        End If
    Next n
End Sub

' follow a GO TO target once during the context walk so arms and
' assignment sites beyond exit-jumps are still discovered
Private Sub CtxGoTo_(ByVal lbl As String, ByVal stack As Collection, ByVal ctx As Collection)
    If Left$(lbl, 6) <> "GO TO " Then Exit Sub
    Dim tgt As String
    tgt = Mid$(lbl, 7)
    If InStr(tgt, " ") > 0 Then Exit Sub
    If mTermSecs.Exists(tgt) Then Exit Sub   ' terminator bodies stay out
    If mCtxVisited.Exists(tgt) Then Exit Sub
    If OnStack_(stack, tgt) Then Exit Sub
    Dim ox As OrderedDict
    Set ox = OwnerByName_(tgt)
    If ox Is Nothing Then Exit Sub
    mCtxVisited.Add tgt, True
    stack.Add tgt
    CtxWalk_ RangeNodes_(CLng(ox.Item("line")), CLng(ox.Item("secEnd"))), stack, ctx
    stack.Remove stack.Count
End Sub

' parse "PERFORM tgt [THRU y] [loop tail]" - True when the body should be
' inlined. Loop forms (VARYING / UNTIL / n TIMES) are walked ONCE: the
' standard static approximation that makes loop-body branches reachable
' (previously they were skipped entirely, leaving every arm inside
' loop-performed paragraphs uncovered).
Private Function ParsePerform_(ByVal lbl As String, ByRef tgt As String, ByRef thru As String) As Boolean
    ParsePerform_ = False
    If Left$(lbl, 8) <> "PERFORM " Then Exit Function
    Dim rest As String, p As Long, tail As String
    rest = Mid$(lbl, 9)
    thru = ""
    p = InStr(rest, " ")
    If p = 0 Then
        tgt = rest
        ParsePerform_ = True
        Exit Function
    End If
    tgt = Left$(rest, p - 1)
    If tgt Like "*[!A-Z0-9-]*" Then Exit Function   ' not a procedure name
    tail = Mid$(rest, p + 1)
    If Left$(tail, 5) = "THRU " Then
        thru = Mid$(tail, 6)
    ElseIf Left$(tail, 8) = "THROUGH " Then
        thru = Mid$(tail, 9)
    ElseIf IsLoopTail_(tail) Then
        ParsePerform_ = True
        Exit Function
    Else
        Exit Function   ' inline PERFORM block etc: stays a generic action
    End If
    p = InStr(thru, " ")
    If p > 0 Then
        tail = Mid$(thru, p + 1)
        thru = Left$(thru, p - 1)
        If Not IsLoopTail_(tail) Then Exit Function
    End If
    ParsePerform_ = True
End Function

Private Function IsLoopTail_(ByVal tail As String) As Boolean
    IsLoopTail_ = False
    If Left$(tail, 8) = "VARYING " Or Left$(tail, 6) = "UNTIL " Then
        IsLoopTail_ = True
        Exit Function
    End If
    ' COBOL85 WITH TEST BEFORE/AFTER prefix wraps a loop tail
    If Left$(tail, 17) = "WITH TEST BEFORE " Then
        IsLoopTail_ = IsLoopTail_(Mid$(tail, 18))
        Exit Function
    End If
    If Left$(tail, 16) = "WITH TEST AFTER " Then
        IsLoopTail_ = IsLoopTail_(Mid$(tail, 17))
        Exit Function
    End If
    ' "<digits> TIMES" exactly. A bare "TIMES" tail means the token BEFORE
    ' it was the repeat count (inline count form) - not inlinable.
    Dim p As Long, cnt As String
    p = InStr(tail, " ")
    If p > 0 Then
        cnt = Left$(tail, p - 1)
        If Mid$(tail, p + 1) = "TIMES" And Len(cnt) > 0 Then
            If Not cnt Like "*[!0-9]*" Then IsLoopTail_ = True
        End If
    End If
End Function

' inline-expand a PERFORM (plain / THRU / loop form) during the context walk
Private Sub CtxPerform_(ByVal lbl As String, ByVal stack As Collection, ByVal ctx As Collection)
    Dim tgt As String, thru As String
    If Not ParsePerform_(lbl, tgt, thru) Then Exit Sub
    If mTermSecs.Exists(tgt) Then Exit Sub
    If OnStack_(stack, tgt) Then Exit Sub
    If mCtxVisited.Exists(tgt) Then Exit Sub
    Dim ox As OrderedDict
    Set ox = OwnerByName_(tgt)
    If ox Is Nothing Then Exit Sub
    mCtxVisited.Add tgt, True
    Dim hi As Long
    If CStr(ox.Item("kind")) = "section" Then hi = CLng(ox.Item("secEnd")) Else hi = CLng(ox.Item("ownerEnd"))
    If Len(thru) > 0 Then
        Dim oy As OrderedDict
        Set oy = OwnerByName_(thru)
        If Not oy Is Nothing Then
            If CStr(oy.Item("kind")) = "section" Then hi = CLng(oy.Item("secEnd")) Else hi = CLng(oy.Item("ownerEnd"))
        End If
    End If
    stack.Add tgt
    CtxWalk_ RangeNodes_(CLng(ox.Item("line")), hi), stack, ctx
    stack.Remove stack.Count
End Sub

Private Sub RegisterCtx_(ByVal token As String, ByVal ctx As Collection)
    If mArmCtx.Exists(token) Then Exit Sub   ' first reaching context wins
    Dim c As Collection, v As Variant
    Set c = New Collection
    For Each v In ctx
        c.Add CStr(v)
    Next v
    mArmCtx.Add token, c
End Sub

Private Function CtxPlus_(ByVal ctx As Collection, ByVal token As String) As Collection
    Dim c As Collection, v As Variant
    Set c = New Collection
    For Each v In ctx
        c.Add CStr(v)
    Next v
    c.Add token
    Set CtxPlus_ = c
End Function

' default-arm preference: avoid the arm that runs straight into an ABEND
' terminator (top-level PERFORM of a registered/auto-detected terminator)
' True if every path through a SECTION reaches a terminator: a top-level
' PERFORM of a known terminator (or GOBACK / STOP RUN / EXIT PROGRAM) appears
' before any top-level branch or jump-away. Sound for the common straight-line
' "always-abend" error handler; returns False on a top-level branch (it might
' fall through and return normally).
Private Function SectionAlwaysTerminates_(ByVal lo As Long, ByVal hi As Long) As Boolean
    SectionAlwaysTerminates_ = False
    Dim nodes As Collection, n As OrderedDict, t As String, lbl As String, tgt As String, thru As String
    Set nodes = RangeNodes_(lo, hi)
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "if" Or t = "evaluate" Or t = "search" Then
            Exit Function   ' top-level branch - may return normally
        ElseIf t = "action" Then
            lbl = CStr(n.Item("label"))
            If lbl = "GOBACK" Or Left$(lbl, 8) = "STOP RUN" Or Left$(lbl, 12) = "EXIT PROGRAM" Then
                SectionAlwaysTerminates_ = True
                Exit Function
            End If
            If Left$(lbl, 6) = "GO TO " Then
                tgt = Trim$(Mid$(lbl, 7))
                If InStr(tgt, " ") = 0 Then
                    If mTermSecs.Exists(tgt) Then SectionAlwaysTerminates_ = True
                End If
                Exit Function   ' jump away - stop scanning either way
            End If
            If ParsePerform_(lbl, tgt, thru) Then
                If mTermSecs.Exists(tgt) Then
                    SectionAlwaysTerminates_ = True
                    Exit Function
                End If
            End If
        End If
    Next n
End Function

Private Function BlockAbends_(ByVal list As Collection) As Boolean
    Dim n As OrderedDict, lbl As String
    BlockAbends_ = False
    For Each n In list
        If CStr(n.Item("type")) = "action" Then
            lbl = CStr(n.Item("label"))
            If Left$(lbl, 8) = "PERFORM " Then
                If mTermSecs.Exists(Mid$(lbl, 9)) Then
                    BlockAbends_ = True
                    Exit Function
                End If
            End If
        End If
    Next n
End Function

'======================================================================
' Structure helpers (owners / cut set / group maps / condition items)
'======================================================================
Private Sub BuildOwners_(ByVal struct As OrderedDict)
    Dim raw As Collection
    Set raw = New Collection
    Dim e As OrderedDict, o As OrderedDict
    For Each e In struct.Item("sections")
        Set o = New OrderedDict
        o.Add "name", UCase$(CStr(e.Item("name")))
        o.Add "line", CLng(e.Item("line"))
        o.Add "kind", "section"
        raw.Add o
    Next e
    For Each e In struct.Item("paragraphs")
        Set o = New OrderedDict
        o.Add "name", UCase$(CStr(e.Item("name")))
        o.Add "line", CLng(e.Item("line"))
        o.Add "kind", "para"
        raw.Add o
    Next e

    ' selection sort by line
    Set mOwners = New Collection
    Dim used() As Boolean, n As Long, k As Long, i As Long, best As Long, bestLine As Long
    n = raw.Count
    If n = 0 Then Exit Sub
    ReDim used(1 To n)
    For k = 1 To n
        best = 0: bestLine = 2000000000
        For i = 1 To n
            If Not used(i) Then
                If CLng(raw(i).Item("line")) < bestLine Then
                    bestLine = CLng(raw(i).Item("line")): best = i
                End If
            End If
        Next i
        used(best) = True
        mOwners.Add raw(best)
    Next k

    Dim oe As Long, se As Long, j As Long
    For i = 1 To mOwners.Count
        Set o = mOwners(i)
        oe = 2000000000
        If i < mOwners.Count Then oe = CLng(mOwners(i + 1).Item("line")) - 1
        se = 2000000000
        For j = i + 1 To mOwners.Count
            If CStr(mOwners(j).Item("kind")) = "section" Then
                se = CLng(mOwners(j).Item("line")) - 1
                Exit For
            End If
        Next j
        o.Add "ownerEnd", oe
        o.Add "secEnd", se
    Next i
End Sub

' a literal MOVE (MOVE ZERO/literal TO f1 f2 ...) returns its target list f1 f2
' ...; empty when the source is not a single literal. Used to spot the DB
' response-code pre-clear that sits right before an EXEC.
Private Function LiteralMoveTargets_(ByVal lbl As String) As String
    LiteralMoveTargets_ = ""
    If Left$(lbl, 5) <> "MOVE " Then Exit Function
    Dim p As Long, src As String, lit As String
    p = InStr(lbl, " TO ")
    If p = 0 Then Exit Function
    src = Trim$(Mid$(lbl, 6, p - 6))
    If Not GetLiteral_(src, lit) Then Exit Function
    LiteralMoveTargets_ = Trim$(Mid$(lbl, p + 4))
End Function

' map each EXEC ... END-EXEC node to the fields cleared by the literal MOVE
' immediately before it (the DB response-code pre-clear idiom). The walk forgets
' those fields' literal at the EXEC so a later branch on the DB result is
' steerable. An EXEC with no preceding literal MOVE (field used directly) is NOT
' recorded, so such sections behave exactly as before.
Private Sub BuildExecClears_(ByVal nodes As Collection)
    Dim i As Long, n As OrderedDict, prev As OrderedDict, t As String, lbl As String, cl As String
    For i = 1 To nodes.Count
        Set n = nodes(i)
        t = CStr(n.Item("type"))
        If t = "action" Then
            lbl = CStr(n.Item("label"))
            If Left$(lbl, 5) = "EXEC " And i > 1 Then
                Set prev = nodes(i - 1)
                If CStr(prev.Item("type")) = "action" Then
                    cl = LiteralMoveTargets_(CStr(prev.Item("label")))
                    If Len(cl) > 0 Then
                        If Not mExecClears.Exists(CStr(n.Item("startLine"))) Then mExecClears.Add CStr(n.Item("startLine")), cl
                    End If
                End If
            End If
        ElseIf t = "if" Then
            BuildExecClears_ n.Item("thenChildren")
            BuildExecClears_ n.Item("elseChildren")
        ElseIf t = "evaluate" Then
            BuildExecClears_ n.Item("cases")
        ElseIf t = "search" Then
            BuildExecClears_ n.Item("atEndChildren")
            BuildExecClears_ n.Item("cases")
        ElseIf t = "when" Then
            BuildExecClears_ n.Item("children")
        End If
    Next i
End Sub

Private Sub BuildCut_()
    Set mCut = New OrderedDict
    Set mGotoCut = New OrderedDict
    CollectPerformTargets_ mNodes
End Sub

Private Sub CollectPerformTargets_(ByVal nodes As Collection)
    Dim n As OrderedDict, t As String
    Dim tgt As String, thru As String
    Dim lbl As String, gt As String, gpos As Long
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "action" Then
            ' all inlinable forms (plain / THRU / loop) cut the fall-through
            ' range, or the body would run twice on the same trace
            If ParsePerform_(CStr(n.Item("label")), tgt, thru) Then
                If Not mCut.Exists(tgt) Then mCut.Add tgt, True
            End If
            ' GO TO targets: collected separately (orphan-root exclusion only,
            ' never fed to mCut - that would wrongly cap walk ranges at them)
            lbl = CStr(n.Item("label"))
            If Left$(lbl, 6) = "GO TO " Then
                gt = Trim$(Mid$(lbl, 7))
                gpos = InStr(gt, " ")
                If gpos > 0 Then gt = Left$(gt, gpos - 1)
                If Len(gt) > 0 Then
                    If Not mGotoCut.Exists(gt) Then mGotoCut.Add gt, True
                End If
            End If
        ElseIf t = "if" Then
            CollectPerformTargets_ n.Item("thenChildren")
            CollectPerformTargets_ n.Item("elseChildren")
        ElseIf t = "evaluate" Then
            CollectPerformTargets_ n.Item("cases")
        ElseIf t = "search" Then
            CollectPerformTargets_ n.Item("atEndChildren")
            CollectPerformTargets_ n.Item("cases")
        ElseIf t = "when" Then
            CollectPerformTargets_ n.Item("children")
        End If
    Next n
End Sub

' parent/descendant maps from the DATA DIVISION level hierarchy
Private Sub BuildGroupMaps_(ByVal lines As Collection)
    Set mDesc = New OrderedDict
    Set mAnc = New OrderedDict
    Dim items As Collection
    Set items = CobolData.Get_DataItems(lines)
    Dim stk As Collection
    Set stk = New Collection
    Dim it As OrderedDict, lv As Long, nm As String, lvS As String
    For Each it In items
        ' COPY entries carry no level (empty string) - the copybook body is
        ' external, so they contribute nothing to the group hierarchy. Skip
        ' anything whose level is not numeric (CLng("") raises type mismatch
        ' - this killed Analyze_Flow on real sources full of COPY lines).
        lvS = CStr(it.Item("level"))
        If Len(lvS) = 0 Or Not IsNumeric(lvS) Then GoTo NextItem
        lv = CLng(lvS)
        nm = UCase$(CStr(it.Item("name")))
        Do While stk.Count > 0
            If CLng(stk(stk.Count).Item("lv")) >= lv Then
                stk.Remove stk.Count
            Else
                Exit Do
            End If
        Loop
        Dim anc As Collection
        Set anc = New Collection
        Dim s As OrderedDict
        For Each s In stk
            AddToListMap_ mDesc, CStr(s.Item("nm")), nm
            anc.Add CStr(s.Item("nm"))
        Next s
        If Not mAnc.Exists(nm) Then mAnc.Add nm, anc
        Dim fr As OrderedDict
        Set fr = New OrderedDict
        fr.Add "lv", lv
        fr.Add "nm", nm
        stk.Add fr
NextItem:
    Next it
End Sub

Private Sub AddToListMap_(ByVal map As OrderedDict, ByVal key As String, ByVal value As String)
    If Not map.Exists(key) Then map.Add key, New Collection
    map.Item(key).Add value
End Sub

' identifiers appearing in any branch condition (for key-assign flagging)
Private Sub BuildCondItems_(ByVal nodes As Collection)
    Set mCondItems = New OrderedDict
    CollectCondItems_ nodes
End Sub

Private Sub CollectCondItems_(ByVal nodes As Collection)
    Dim n As OrderedDict, t As String
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "if" Then
            AddCondIdents_ CStr(n.Item("condition"))
            CollectCondItems_ n.Item("thenChildren")
            CollectCondItems_ n.Item("elseChildren")
        ElseIf t = "evaluate" Then
            AddCondIdents_ CStr(n.Item("expression"))
            CollectCondItems_ n.Item("cases")
        ElseIf t = "search" Then
            CollectCondItems_ n.Item("atEndChildren")
            CollectCondItems_ n.Item("cases")
        ElseIf t = "when" Then
            AddCondIdents_ CStr(n.Item("condition"))
            CollectCondItems_ n.Item("children")
        End If
    Next n
End Sub

Private Sub AddCondIdents_(ByVal cond As String)
    Static rxId As Object
    If rxId Is Nothing Then
        Set rxId = CreateObject("VBScript.RegExp")
        rxId.Pattern = "[A-Z][A-Z0-9-]+"
        rxId.Global = True
        rxId.IgnoreCase = False
    End If
    Dim m As Object, i As Long, w As String
    Set m = rxId.Execute(cond)
    For i = 0 To m.Count - 1
        w = m.Item(i).Value
        Select Case w
            Case "OR", "AND", "NOT", "ZERO", "ZEROS", "ZEROES", "SPACE", "SPACES", "OTHER", "THEN"
            Case Else
                If Not mCondItems.Exists(w) Then mCondItems.Add w, True
        End Select
    Next i
End Sub

Private Function BuildSecLabels_(ByVal norm As OrderedDict) As OrderedDict
    Dim labels As OrderedDict
    Set labels = New OrderedDict
    Set BuildSecLabels_ = labels
    If Not norm.Exists("Comments") Then Exit Function
    Dim comments As Collection
    Set comments = norm.Item("Comments")
    Dim po As String, pc As String
    po = ChrW$(&HFF08)   ' full-width (
    pc = ChrW$(&HFF09)   ' full-width )
    Dim o As OrderedDict, c As OrderedDict, secLine As Long, txt As String, p1 As Long, p2 As Long
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" Then
            secLine = CLng(o.Item("line"))
            For Each c In comments
                If CLng(c.Item("Number")) >= secLine - 4 And CLng(c.Item("Number")) < secLine Then
                    txt = CStr(c.Item("Text"))
                    p1 = InStr(txt, po)
                    If p1 > 0 Then
                        p2 = InStr(p1 + 1, txt, pc)
                        If p2 > p1 Then
                            If Not labels.Exists(CStr(o.Item("name"))) Then
                                labels.Add CStr(o.Item("name")), Mid$(txt, p1 + 1, p2 - p1 - 1)
                            End If
                        End If
                    End If
                End If
            Next c
        End If
    Next o
End Function

' Entry = the procedure section owning the first statement AT OR AFTER the
' first non-data section. Anchoring on that section keeps stray pre-procedure
' parse artifacts (DATA DIVISION quirks) from breaking entry detection.
Private Function FindEntry_() As OrderedDict
    Dim o As OrderedDict, firstProc As OrderedDict, found As OrderedDict
    Set firstProc = Nothing
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" And Not IsDataSection_(CStr(o.Item("name"))) Then
            Set firstProc = o
            Exit For
        End If
    Next o
    Set FindEntry_ = Nothing
    If firstProc Is Nothing Then Exit Function

    Dim firstStmt As Long, n As OrderedDict
    firstStmt = 2000000000
    For Each n In mNodes
        If CLng(n.Item("startLine")) >= CLng(firstProc.Item("line")) Then
            If CLng(n.Item("startLine")) < firstStmt Then firstStmt = CLng(n.Item("startLine"))
        End If
    Next n

    Set found = firstProc
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" And CLng(o.Item("line")) <= firstStmt Then
            If Not IsDataSection_(CStr(o.Item("name"))) Then Set found = o
        End If
    Next o
    Set FindEntry_ = found
End Function

' True when [lo, hi] contains an explicit normal terminator (GOBACK / STOP RUN
' / EXIT PROGRAM). When FALSE for the program's main entry, falling off the
' end is an implicit normal return, so RunWalk_ promotes a Term="" top-level
' walk to "goback" (else a bare-EXIT-ending program yields no normal cases and
' every arm reads Œo˜H\’z•s‰Â).
Private Function LinesHaveNormalTerm_(ByVal lines As Collection, ByVal lo As Long, ByVal hi As Long) As Boolean
    LinesHaveNormalTerm_ = False
    Dim le As OrderedDict, t As String, ln As Long
    For Each le In lines
        ln = CLng(le.Item("Number"))
        If ln >= lo And ln <= hi Then
            t = UCase$(Trim$(CStr(le.Item("Text"))))
            If t = "GOBACK" Or t = "GOBACK." _
               Or Left$(t, 8) = "STOP RUN" Or Left$(t, 12) = "EXIT PROGRAM" Then
                LinesHaveNormalTerm_ = True
                Exit Function
            End If
        End If
    Next le
End Function

' a unit/orphan walk that ran to the section's EXIT carries Term="" (no
' GOBACK); for a SECTION under unit test that IS a normal return, so promote
' it before AddCand_ (which would otherwise drop a terminator-less trace).
Private Sub AddUnitCand_(ByVal cands As Collection, ByVal covered As OrderedDict, ByVal w As OrderedDict)
    If w Is Nothing Then Exit Sub
    If CStr(w.Item("Term")) = "" Then w.Add "Term", "goback"
    AddCand_ cands, covered, w
End Sub

' procedure SECTIONs the main flow never PERFORMs - candidate unit-test entry
' points. Excludes the entry itself, data-division sections, terminators, and
' PERFORM targets (mCut); the last are reached transitively, so only true roots
' are returned. The caller still gates on SecHasUncoveredArm_.
Private Function OrphanRoots_(ByVal entrySec As OrderedDict) As Collection
    Dim res As Collection, o As OrderedDict, nm As String
    Set res = New Collection
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" Then
            nm = CStr(o.Item("name"))
            If Not IsDataSection_(nm) _
               And nm <> CStr(entrySec.Item("name")) _
               And Not mCut.Exists(nm) _
               And Not mGotoCut.Exists(nm) _
               And Not mTermSecs.Exists(nm) Then
                res.Add o
            End If
        End If
    Next o
    Set OrphanRoots_ = res
End Function

' True when section [line, secEnd] owns at least one still-uncovered arm.
Private Function SecHasUncoveredArm_(ByVal oSec As OrderedDict, ByVal arms As Collection, _
                                     ByVal covered As OrderedDict) As Boolean
    SecHasUncoveredArm_ = False
    Dim lo As Long, hi As Long, a As OrderedDict, ln As Long
    lo = CLng(oSec.Item("line"))
    hi = CLng(oSec.Item("secEnd"))
    For Each a In arms
        ln = CLng(a.Item("Line"))
        If ln >= lo And ln <= hi Then
            If Not covered.Exists(CStr(a.Item("Token"))) Then
                SecHasUncoveredArm_ = True
                Exit Function
            End If
        End If
    Next a
End Function

Private Function IsDataSection_(ByVal nm As String) As Boolean
    Select Case nm
        Case "WORKING-STORAGE", "LOCAL-STORAGE", "LINKAGE", "FILE", _
             "CONFIGURATION", "INPUT-OUTPUT", "SPECIAL-NAMES", "FILE-CONTROL", _
             "I-O-CONTROL", "COMMUNICATION", "REPORT", "SCREEN"
            IsDataSection_ = True
        Case Else
            IsDataSection_ = False
    End Select
End Function

' the comment just above a section header, used as the section's kanji
' label. Scans a small window upward: decoration rows (****) are skipped,
' blank lines (gaps in the line-number space) are skipped, and the scan
' stops at the first real CODE line so a previous section's tail comment
' is never picked up.
Private Function SecNote_(ByVal cmtMap As OrderedDict, ByVal codeSet As OrderedDict, _
                          ByVal lineNo As Long) As String
    SecNote_ = ""
    Dim k As Long, ctext As String, bare As String
    For k = lineNo - 1 To lineNo - 6 Step -1
        If k < 1 Then Exit For
        If cmtMap.Exists(CStr(k)) Then
            ctext = CStr(cmtMap.Item(CStr(k)))
            bare = Replace(Replace(Replace(Replace(ctext, "*", ""), "-", ""), "=", ""), "/", "")
            bare = Replace(bare, " ", "")
            If Len(bare) > 0 Then
                SecNote_ = Trim$(Replace(ctext, "*", ""))
                Exit Function
            End If
            ' decoration-only comment - keep scanning upward
        ElseIf codeSet.Exists(CStr(k)) Then
            Exit For   ' hit a real code line above - stop
        End If
        ' else: blank line (not comment, not code) - skip and keep scanning
    Next k
End Function

Private Function OwnerByName_(ByVal nm As String) As OrderedDict
    ' indexed lookup - this is called per trace per PERFORM
    If Not mOwnerIdx Is Nothing Then
        If mOwnerIdx.Exists(nm) Then
            Set OwnerByName_ = mOwnerIdx.Item(nm)
        Else
            Set OwnerByName_ = Nothing
        End If
        Exit Function
    End If
    Dim o As OrderedDict
    For Each o In mOwners
        If CStr(o.Item("name")) = nm Then
            Set OwnerByName_ = o
            Exit Function
        End If
    Next o
    Set OwnerByName_ = Nothing
End Function

Private Function CapHi_(ByVal lo As Long, ByVal hi As Long) As Long
    Dim o As OrderedDict, c As Long, ln As Long
    c = 2000000000
    For Each o In mOwners
        If mCut.Exists(CStr(o.Item("name"))) Then
            ln = CLng(o.Item("line"))
            If ln > lo And ln <= hi And ln < c Then c = ln
        End If
    Next o
    If c < 2000000000 Then CapHi_ = c - 1 Else CapHi_ = hi
End Function

Private Function NodesInRange_(ByVal lo As Long, ByVal hi As Long) As Collection
    Dim c As Collection, n As OrderedDict, ln As Long
    Set c = New Collection
    For Each n In mNodes
        ln = CLng(n.Item("startLine"))
        If ln >= lo And ln <= hi Then c.Add n
    Next n
    Set NodesInRange_ = c
End Function

'======================================================================
' Trace plumbing
'======================================================================
' Traces hold their Arms / Events / Calls as immutable ConsList heads
' (Nothing = empty). Appending and cloning are O(1); the lists are
' materialized into ordered Collections only for the finally-selected
' cases (BuildCase_ / SelectCases_). This removed the per-fork
' full-Collection copies that froze Excel on 1000+ line programs.
Private Function NewTrace_() As OrderedDict
    Dim t As OrderedDict
    Set t = New OrderedDict
    t.Add "Arms", Nothing
    t.Add "Events", Nothing
    t.Add "Calls", Nothing
    t.Add "Env", New OrderedDict
    t.Add "Term", ""
    t.Add "TriggerLine", 0
    Set NewTrace_ = t
End Function

' (CloneTrace_ removed: directed walks never fork, so traces are mutated in
'  place and synthesized-failure prefixes are produced by a stop-at-call
'  re-walk instead of snapshot cloning.)

' O(1) list append: new head referencing the previous one.
Private Function Cons_(ByVal head As ConsList, ByVal item As Variant) As ConsList
    Dim n As ConsList
    Set n = New ConsList
    If IsObject(item) Then
        Set n.V = item
    Else
        n.V = item
    End If
    Set n.Prev = head
    If head Is Nothing Then n.N = 1 Else n.N = head.N + 1
    If mConsReg Is Nothing Then Set mConsReg = New Collection
    mConsReg.Add n
    Set Cons_ = n
End Function

' Iteratively break every Prev link so node teardown never recurses.
' Called once per run after all selected cases are materialized.
Private Sub UnlinkCons_()
    If mConsReg Is Nothing Then Exit Sub
    Dim n As ConsList
    For Each n In mConsReg
        Set n.Prev = Nothing
    Next n
    Set mConsReg = New Collection
End Sub

Private Function ConsCount_(ByVal head As ConsList) As Long
    If head Is Nothing Then ConsCount_ = 0 Else ConsCount_ = head.N
End Function

' Materialize a cons list into a Collection in original (append) order.
Private Function ConsToList_(ByVal head As ConsList) As Collection
    Dim c As Collection
    Set c = New Collection
    Set ConsToList_ = c
    Dim n As Long
    n = ConsCount_(head)
    If n = 0 Then Exit Function
    Dim arr() As Variant, cur As ConsList, i As Long
    ReDim arr(1 To n)
    Set cur = head
    For i = n To 1 Step -1
        If IsObject(cur.V) Then Set arr(i) = cur.V Else arr(i) = cur.V
        Set cur = cur.Prev
    Next i
    For i = 1 To n
        c.Add arr(i)
    Next i
End Function

Private Sub AddEvent_(ByVal t As OrderedDict, ByVal kind As String, ByVal text As String, ByVal lineNo As Long)
    If Not mRecordEvents Then Exit Sub   ' pass 1 needs arms/term only
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "Kind", kind
    e.Add "Text", text
    e.Add "Line", lineNo
    t.Add "Events", Cons_(t.Item("Events"), e)
End Sub

Private Sub AddEnterEvent_(ByVal t As OrderedDict, ByVal secName As String, ByVal secLabels As OrderedDict, ByVal lineNo As Long)
    If Not mRecordEvents Then Exit Sub
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "Kind", "enter"
    e.Add "Text", secName
    If secLabels.Exists(secName) Then
        e.Add "Label", CStr(secLabels.Item(secName))
    Else
        e.Add "Label", ""
    End If
    e.Add "Line", lineNo
    t.Add "Events", Cons_(t.Item("Events"), e)
End Sub

Private Sub AddArmEvent_(ByVal t As OrderedDict, ByVal token As String, ByVal armName As String, _
                         ByVal cond As String, ByVal lineNo As Long)
    t.Add "Arms", Cons_(t.Item("Arms"), token)   ' arms always (selection needs them)
    If Not mRecordEvents Then Exit Sub
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "Kind", "arm"
    e.Add "Token", token
    e.Add "Arm", armName
    e.Add "Cond", cond
    e.Add "Text", armName & " : " & cond
    e.Add "Line", lineNo
    t.Add "Events", Cons_(t.Item("Events"), e)
End Sub

Private Sub AddAssignEvent_(ByVal t As OrderedDict, ByVal dst As String, ByVal srcTxt As String, _
                            ByVal kind As String, ByVal lineNo As Long)
    If Not mRecordEvents Then Exit Sub
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "Kind", "assign"
    e.Add "Dst", dst
    e.Add "Src", srcTxt
    e.Add "AKind", kind
    e.Add "IsKey", mCondItems.Exists(dst)
    e.Add "Text", dst & " <- " & srcTxt
    e.Add "Line", lineNo
    t.Add "Events", Cons_(t.Item("Events"), e)
End Sub

'======================================================================
' Literal environment + feasibility
'======================================================================
Private Function GetLiteral_(ByVal srcTxt As String, ByRef out As String) As Boolean
    Dim s As String
    s = Trim$(srcTxt)
    GetLiteral_ = False
    If Len(s) >= 2 Then
        If Left$(s, 1) = "'" And Right$(s, 1) = "'" Then
            out = Mid$(s, 2, Len(s) - 2)
            ' a quote INSIDE means this was a compound expression
            ' ("'1' OR X = '2'"), not a single literal - refuse it
            If InStr(out, "'") > 0 Then Exit Function
            GetLiteral_ = True
            Exit Function
        End If
    End If
    If s Like String(Len(s), "#") And Len(s) > 0 Then
        out = NormDigits_(s)
        GetLiteral_ = True
        Exit Function
    End If
    Select Case s
        Case "ZERO", "ZEROS", "ZEROES"
            out = "0": GetLiteral_ = True
        Case "SPACE", "SPACES"
            out = " ": GetLiteral_ = True
    End Select
End Function

Private Sub Invalidate_(ByVal env As OrderedDict, ByVal item As String)
    RemoveKey_ env, item
    Dim v As Variant
    If mDesc.Exists(item) Then
        For Each v In mDesc.Item(item): RemoveKey_ env, CStr(v): Next v
    End If
    If mAnc.Exists(item) Then
        For Each v In mAnc.Item(item): RemoveKey_ env, CStr(v): Next v
    End If
End Sub

' OrderedDict has no Remove - rebuild without the key when present.
Private Sub RemoveKey_(ByVal env As OrderedDict, ByVal key As String)
    If Not env.Exists(key) Then Exit Sub
    Dim ks As Collection, v As Variant, tmp As OrderedDict
    Set ks = env.Keys
    Set tmp = New OrderedDict
    For Each v In ks
        If CStr(v) <> key Then tmp.Add CStr(v), env.Item(CStr(v))
    Next v
    ' copy back
    Dim ks2 As Collection
    Set ks2 = tmp.Keys
    ClearDict_ env
    For Each v In ks2
        env.Add CStr(v), tmp.Item(CStr(v))
    Next v
End Sub

' OrderedDict lacks Clear too: emulate by re-Add semantics is impossible,
' so mark cleared via sentinel rebuild - instead we swap contents using a
' helper that relies on Add-overwrite. Simplest correct approach: rebuild
' by overwriting every key with Empty is wrong; we need true clear. We use
' a fresh dict in callers where possible; here we shrink by overwriting:
Private Sub ClearDict_(ByVal d As OrderedDict)
    ' OrderedDict has no Clear; emulate by removing through reconstruction
    ' is not possible in place. Callers of RemoveKey_ accept that cleared
    ' keys are set to Empty instead (Exists still true but value Empty).
    Dim ks As Collection, v As Variant
    Set ks = d.Keys
    For Each v In ks
        d.Add CStr(v), Empty
    Next v
End Sub

Private Function EnvHas_(ByVal env As OrderedDict, ByVal item As String) As Boolean
    EnvHas_ = False
    If Not env.Exists(item) Then Exit Function
    If IsEmpty(env.Item(item)) Then Exit Function
    EnvHas_ = True
End Function

Private Function NormVal_(ByVal v As String) As String
    If v Like String(Len(v), "#") And Len(v) > 0 Then
        NormVal_ = NormDigits_(v)
    Else
        NormVal_ = v
    End If
End Function

' Digit-string normalization without Double precision loss / overflow:
' strip leading zeros manually for long values, CDbl only for short ones.
Private Function NormDigits_(ByVal s As String) As String
    If Len(s) <= 15 Then
        NormDigits_ = CStr(CDbl(s))
    Else
        Dim i As Long
        i = 1
        Do While i < Len(s) And Mid$(s, i, 1) = "0"
            i = i + 1
        Loop
        NormDigits_ = Mid$(s, i)
    End If
End Function

' Feasibility of (THEN, ELSE) for a condition under env. Conservative: prune
' only when the condition is equality-style on a single known item.
Private Sub TestFeasible_(ByVal cond As String, ByVal env As OrderedDict, _
                          ByRef okThen As Boolean, ByRef okElse As Boolean)
    okThen = True
    okElse = True
    Dim c As String
    c = Trim$(cond)
    If InStr(c, " AND ") > 0 Then Exit Sub

    Static rxEq As Object, rxNe As Object
    If rxEq Is Nothing Then
        Set rxEq = CreateObject("VBScript.RegExp")
        rxEq.Pattern = "^([A-Z][A-Z0-9-]*)\s*=\s*(.+)$"
        rxEq.IgnoreCase = False
        Set rxNe = CreateObject("VBScript.RegExp")
        rxNe.Pattern = "^([A-Z][A-Z0-9-]*)\s+NOT\s*=\s*(.+)$"
        rxNe.IgnoreCase = False
    End If

    Dim m As Object, lit As String
    ' NOT-equal single term (only when no OR joins further clauses -
    ' otherwise the regex tail would swallow the rest of the condition)
    If InStr(c, " OR ") = 0 Then
        Set m = rxNe.Execute(c)
        If m.Count > 0 Then
            If GetLiteral_(m.Item(0).SubMatches(1), lit) Then
                Dim itemN As String
                itemN = m.Item(0).SubMatches(0)
                If EnvHas_(env, itemN) Then
                    If NormVal_(CStr(env.Item(itemN))) = NormVal_(lit) Then
                        okThen = False
                    Else
                        okElse = False
                    End If
                End If
            End If
            Exit Sub
        End If
    ElseIf InStr(c, " NOT ") > 0 Then
        ' mixed NOT with OR: too complex - never prune
        Exit Sub
    End If

    ' OR-joined equalities on the same item
    Dim parts() As String, i As Long, itemNm As String, vals As Collection
    parts = Split(c, " OR ")
    Set vals = New Collection
    itemNm = ""
    For i = LBound(parts) To UBound(parts)
        Set m = rxEq.Execute(Trim$(parts(i)))
        If m.Count = 0 Then Exit Sub
        If Not GetLiteral_(m.Item(0).SubMatches(1), lit) Then Exit Sub
        If itemNm = "" Then
            itemNm = m.Item(0).SubMatches(0)
        ElseIf itemNm <> m.Item(0).SubMatches(0) Then
            Exit Sub
        End If
        vals.Add NormVal_(lit)
    Next i
    If itemNm = "" Then Exit Sub
    If Not EnvHas_(env, itemNm) Then Exit Sub

    Dim known As String, hit As Boolean, vv As Variant
    known = NormVal_(CStr(env.Item(itemNm)))
    hit = False
    For Each vv In vals
        If CStr(vv) = known Then hit = True
    Next vv
    If hit Then okElse = False Else okThen = False
End Sub

'======================================================================
' Enumeration
'======================================================================
Private Function ApplyRange_(ByVal lo As Long, ByVal hi As Long, ByVal stack As Collection, _
                             ByVal traces As Collection, ByVal secLabels As OrderedDict) As Collection
    Set ApplyRange_ = ApplyNodes_(RangeNodes_(lo, hi), stack, traces, secLabels)
End Function

' ranges repeat massively (every walk re-enters the same PERFORM targets):
' resolve + scan once per unique range, then reuse the node list
Private Function RangeNodes_(ByVal lo As Long, ByVal hi As Long) As Collection
    Dim key As String
    key = lo & "|" & hi
    If mRangeCache Is Nothing Then Set mRangeCache = New OrderedDict
    If Not mRangeCache.Exists(key) Then mRangeCache.Add key, NodesInRange_(lo, CapHi_(lo, hi))
    Set RangeNodes_ = mRangeCache.Item(key)
End Function

Private Function ApplyNodes_(ByVal list As Collection, ByVal stack As Collection, _
                             ByVal traces As Collection, ByVal secLabels As OrderedDict) As Collection
    Dim acc As Collection, i As Long
    Set acc = traces
    For i = 1 To list.Count
        Set acc = ApplyOne_(list(i), stack, acc, secLabels)
    Next i
    Set ApplyNodes_ = acc
End Function

Private Function ApplyOne_(ByVal node As OrderedDict, ByVal stack As Collection, _
                           ByVal traces As Collection, ByVal secLabels As OrderedDict) As Collection
    mOps = mOps + 1
    If (mOps And 2047) = 0 Then
        ' visible progress so a long run is distinguishable from a hang.
        ' only during the real pipeline - direct calls (tests) must not
        ' leave a stale status bar behind.
        On Error Resume Next
        If Main.AnalysisBusy() Then Application.StatusBar = Main.STATUS_FLOW & "  (" & mOps & ")"
        On Error GoTo 0
        DoEvents   ' stay responsive on big programs
    End If
    Dim out As Collection
    Set out = New Collection
    Dim tr As OrderedDict, t As String
    t = CStr(node.Item("type"))

    For Each tr In traces
        If CStr(tr.Item("Term")) <> "" Then
            AddCapped_ out, tr
        ElseIf TrSkips_(tr) Then
            AddCapped_ out, tr   ' control jumped away - pass through
        ElseIf t = "action" Then
            ApplyAction_ node, stack, tr, out, secLabels
        ElseIf t = "if" Then
            ' directed: choose exactly one arm (required > default), no fork
            Dim okT As Boolean, okE As Boolean
            TestFeasible_ CStr(node.Item("condition")), tr.Item("Env"), okT, okE
            Dim tkT As String, tkE As String, armSel As String
            tkT = CStr(node.Item("id")) & ":then"
            tkE = CStr(node.Item("id")) & ":else"
            armSel = ""
            ' tier 1: pass-2 replay consumes the recorded arm sequence
            If Not mReplayList Is Nothing Then
                If mReplayIdx <= mReplayList.Count Then
                    Dim rt As String
                    rt = CStr(mReplayList(mReplayIdx))
                    If rt = tkT Then
                        armSel = "T"
                        mReplayIdx = mReplayIdx + 1
                    ElseIf rt = tkE Then
                        armSel = "E"
                        mReplayIdx = mReplayIdx + 1
                    End If
                End If
            End If
            If armSel = "" Then
                If mNeed.Exists(tkT) Then
                    If okT Then
                        armSel = "T"
                    ElseIf mForce Then
                        armSel = "T"          ' force the value-conflicted arm
                    Else
                        mMissed = True
                        mMissCond = CStr(node.Item("condition"))
                        mMissTok = tkT
                        If okE Then armSel = "E"
                    End If
                ElseIf mNeed.Exists(tkE) Then
                    If okE Then
                        armSel = "E"
                    ElseIf mForce Then
                        armSel = "E"          ' force the value-conflicted arm
                    Else
                        mMissed = True
                        mMissCond = CStr(node.Item("condition"))
                        mMissTok = tkE
                        If okT Then armSel = "T"
                    End If
                ElseIf mSweep Then
                    ' tier 3: grab an uncovered non-abend arm, else follow the
                    ' subtree holding more uncovered arms, else default
                    Dim uT As Boolean, uE As Boolean
                    uT = False
                    uE = False
                    If mSweepUncov.Exists(tkT) Then
                        If Not BlockAbends_(node.Item("thenChildren")) Then uT = True
                    End If
                    If mSweepUncov.Exists(tkE) Then
                        If Not BlockAbends_(node.Item("elseChildren")) Then uE = True
                    End If
                    If uT And okT Then
                        armSel = "T"
                    ElseIf uE And okE Then
                        armSel = "E"
                    Else
                        Dim wT As Long, wE As Long
                        wT = 0
                        wE = 0
                        If mSweepW.Exists(tkT) Then wT = CLng(mSweepW.Item(tkT))
                        If mSweepW.Exists(tkE) Then wE = CLng(mSweepW.Item(tkE))
                        If wT > wE And okT Then
                            armSel = "T"
                        ElseIf wE > wT And okE Then
                            armSel = "E"
                        Else
                            armSel = DefaultArm_(node, okT, okE)
                        End If
                    End If
                Else
                    armSel = DefaultArm_(node, okT, okE)
                End If
            End If
            Dim subs As Collection, sb As OrderedDict, seed As Collection
            If armSel = "T" Then
                AddArmEvent_ tr, tkT, "THEN", CStr(node.Item("condition")), CLng(node.Item("startLine"))
                Set seed = New Collection
                seed.Add tr
                Set subs = ApplyNodes_(node.Item("thenChildren"), stack, seed, secLabels)
                For Each sb In subs: AddCapped_ out, sb: Next sb
            ElseIf armSel = "E" Then
                AddArmEvent_ tr, tkE, "ELSE", CStr(node.Item("condition")), CLng(node.Item("startLine"))
                Set seed = New Collection
                seed.Add tr
                Set subs = ApplyNodes_(node.Item("elseChildren"), stack, seed, secLabels)
                For Each sb In subs: AddCapped_ out, sb: Next sb
            Else
                mMissed = True   ' both arms infeasible - the walk dies here
            End If
        ElseIf t = "evaluate" Then
            ApplyEvaluate_ node, stack, tr, out, secLabels
        ElseIf t = "search" Then
            ApplySearch_ node, stack, tr, out, secLabels
        Else
            AddCapped_ out, tr
        End If
    Next tr
    Set ApplyOne_ = out
End Function

' accept a finished candidate walk: materialize its arms once and mark
' them covered (drives the sweep rounds and the fallback skip)
Private Sub AddCand_(ByVal cands As Collection, ByVal covered As OrderedDict, ByVal w As OrderedDict)
    If CStr(w.Item("Term")) = "" Then Exit Sub
    If Not w.Exists("EntrySec") Then w.Add "EntrySec", mCurEntrySec
    Dim al As Collection, v As Variant
    Set al = ConsToList_(w.Item("Arms"))
    w.Add "ArmsL", al
    For Each v In al
        If Not covered.Exists(CStr(v)) Then covered.Add CStr(v), True
    Next v
    cands.Add w
End Sub

' uncovered set + ancestor weights for one sweep round
Private Function BuildSweepState_(ByVal arms As Collection, ByVal covered As OrderedDict) As Boolean
    Set mSweepUncov = New OrderedDict
    Set mSweepW = New OrderedDict
    Dim a As OrderedDict, v As Variant, tok As String
    For Each a In arms
        tok = CStr(a.Item("Token"))
        If Not covered.Exists(tok) Then
            mSweepUncov.Add tok, True
            If mArmCtx.Exists(tok) Then
                For Each v In mArmCtx.Item(tok)
                    If mSweepW.Exists(CStr(v)) Then
                        mSweepW.Add CStr(v), CLng(mSweepW.Item(CStr(v))) + 1
                    Else
                        mSweepW.Add CStr(v), 1
                    End If
                Next v
            End If
        End If
    Next a
    BuildSweepState_ = (mSweepUncov.Count > 0)
End Function

Private Function SweepGain_(ByVal w As OrderedDict, ByVal covered As OrderedDict) As Long
    Dim v As Variant, g As Long
    g = 0
    For Each v In ConsToList_(w.Item("Arms"))
        If Not covered.Exists(CStr(v)) Then g = g + 1
    Next v
    SweepGain_ = g
End Function

' default arm preference: THEN-first (ELSE-first for the prefElse seed),
' flipped when the preferred arm runs straight into an ABEND terminator
' or jumps away (GO TO) - unsteered walks should keep flowing forward
Private Function DefaultArm_(ByVal node As OrderedDict, ByVal okT As Boolean, ByVal okE As Boolean) As String
    DefaultArm_ = ""
    Dim prefT As Boolean, badT As Boolean, badE As Boolean
    badT = BlockAbends_(node.Item("thenChildren")) Or BlockJumps_(node.Item("thenChildren"))
    badE = BlockAbends_(node.Item("elseChildren")) Or BlockJumps_(node.Item("elseChildren"))
    prefT = Not mPrefElse
    If prefT Then
        If badT And Not badE Then prefT = False
    Else
        If badE And Not badT Then prefT = True
    End If
    If prefT Then
        If okT Then DefaultArm_ = "T" Else If okE Then DefaultArm_ = "E"
    Else
        If okE Then DefaultArm_ = "E" Else If okT Then DefaultArm_ = "T"
    End If
End Function

' does a branch body jump away (top-level GO TO)?
Private Function BlockJumps_(ByVal list As Collection) As Boolean
    Dim n As OrderedDict
    BlockJumps_ = False
    For Each n In list
        If CStr(n.Item("type")) = "action" Then
            If Left$(CStr(n.Item("label")), 6) = "GO TO " Then
                BlockJumps_ = True
                Exit Function
            End If
        End If
    Next n
End Function

' assignment targets of an arithmetic verb (ADD/SUBTRACT/MULTIPLY/DIVIDE):
' tokens after GIVING when present, else after TO / FROM / INTO; REMAINDER
' targets included, ROUNDED keywords skipped
Private Function ArithTargets_(ByVal lbl As String) As Collection
    Dim c As Collection
    Set c = New Collection
    Set ArithTargets_ = c
    Dim p As Long, tail As String
    p = InStr(lbl, " GIVING ")
    If p > 0 Then
        tail = Mid$(lbl, p + 8)
    ElseIf Left$(lbl, 4) = "ADD " Then
        p = InStr(lbl, " TO ")
        If p = 0 Then Exit Function
        tail = Mid$(lbl, p + 4)
    ElseIf Left$(lbl, 9) = "SUBTRACT " Then
        p = InStr(lbl, " FROM ")
        If p = 0 Then Exit Function
        tail = Mid$(lbl, p + 6)
    ElseIf Left$(lbl, 9) = "MULTIPLY " Then
        p = InStr(lbl, " BY ")
        If p = 0 Then Exit Function
        tail = Mid$(lbl, p + 4)
    ElseIf Left$(lbl, 7) = "DIVIDE " Then
        p = InStr(lbl, " INTO ")
        If p = 0 Then Exit Function
        tail = Mid$(lbl, p + 6)
    Else
        Exit Function
    End If
    Dim toks() As String, i As Long, t As String
    toks = Split(Trim$(tail), " ")
    For i = LBound(toks) To UBound(toks)
        t = Trim$(toks(i))
        ' targets always precede a conditional phrase / scope terminator
        If t = "ON" Or t = "NOT" Or Left$(t, 4) = "END-" Then Exit For
        If t = "ROUNDED" Or t = "REMAINDER" Or Len(t) = 0 Then
            ' skip keywords; REMAINDER's operand is collected next round
        ElseIf Not t Like "*[!A-Z0-9-]*" And t Like "*[A-Z]*" Then
            ' identifier = charset-clean AND contains a letter (digit-leading
            ' data-names are legal; pure numerics are literals)
            c.Add t
        End If
    Next i
End Function

Private Sub AddCapped_(ByVal out As Collection, ByVal tr As OrderedDict)
    If out.Count >= MAX_TRACES Then
        mTruncated = True
    Else
        out.Add tr
    End If
End Sub

Private Sub ApplyAction_(ByVal node As OrderedDict, ByVal stack As Collection, ByVal tr As OrderedDict, _
                         ByVal out As Collection, ByVal secLabels As OrderedDict)
    Static rxCall As Object, rxMove As Object
    Static rxComp As Object, rxStr As Object, rxInit As Object, rxAcc As Object
    If rxCall Is Nothing Then
        Set rxCall = CreateObject("VBScript.RegExp"): rxCall.Pattern = "^CALL\s+'([A-Z0-9-]+)'(\s+USING\s+(.+))?$": rxCall.IgnoreCase = False
        Set rxMove = CreateObject("VBScript.RegExp"): rxMove.Pattern = "^MOVE\s+(.+?)\s+TO\s+([A-Z0-9-]+(\s+[A-Z0-9-]+)*)$": rxMove.IgnoreCase = False
        Set rxComp = CreateObject("VBScript.RegExp"): rxComp.Pattern = "^COMPUTE\s+([A-Z0-9-]+)\s*=\s*(.+)$": rxComp.IgnoreCase = False
        Set rxStr = CreateObject("VBScript.RegExp"): rxStr.Pattern = "^STRING\s+(.+)\s+INTO\s+([A-Z0-9-]+)$": rxStr.IgnoreCase = False
        Set rxInit = CreateObject("VBScript.RegExp"): rxInit.Pattern = "^INITIALIZE\s+([A-Z0-9-]+)": rxInit.IgnoreCase = False
        Set rxAcc = CreateObject("VBScript.RegExp"): rxAcc.Pattern = "^ACCEPT\s+([A-Z0-9-]+)": rxAcc.IgnoreCase = False
    End If

    Dim lbl As String, ln As Long, m As Object
    lbl = CStr(node.Item("label"))
    ln = CLng(node.Item("startLine"))

    ' terminators
    If lbl = "GOBACK" Or lbl = "STOP RUN" Or lbl = "EXIT PROGRAM" Then
        tr.Add "Term", "goback"
        tr.Add "TriggerLine", ln
        AddEvent_ tr, "term", lbl, ln
        AddCapped_ out, tr
        Exit Sub
    End If

    ' fast dispatch on the first token: this sub runs per trace per action,
    ' and running every regex against every label dominated big-program
    ' runtime. Only the verbs below need the regex machinery.
    Dim v1 As String
    v1 = Left$(lbl & " ", InStr(lbl & " ", " "))
    Select Case v1
        Case "PERFORM ", "CALL ", "MOVE ", "COMPUTE ", "STRING ", "INITIALIZE ", "ACCEPT ", _
             "GO ", "ADD ", "SUBTRACT ", "MULTIPLY ", "DIVIDE ", "READ "
            ' fall through to the verb handlers below
        Case Else
            If Left$(lbl, 5) = "EXEC " Then
                ' embedded DB/CICS call writes its result fields. If a literal
                ' MOVE (the response-code pre-clear) sits right before this EXEC,
                ' those fields are the DB results - forget their literal so a
                ' later branch on the DB result can take any arm (stub sets it).
                If mExecClears.Exists(CStr(ln)) Then
                    Dim ecF() As String, eci As Long
                    ecF = Split(CStr(mExecClears.Item(CStr(ln))), " ")
                    For eci = LBound(ecF) To UBound(ecF)
                        If Len(Trim$(ecF(eci))) > 0 Then Invalidate_ tr.Item("Env"), UCase$(Trim$(ecF(eci)))
                    Next eci
                End If
                AddEvent_ tr, "exec", lbl, ln
            Else
                AddEvent_ tr, "action", lbl, ln
            End If
            AddCapped_ out, tr
            Exit Sub
    End Select

    If v1 = "PERFORM " Then
        ' plain / THRU / loop form (loop bodies are inlined once)
        Dim pTgt As String, pThru As String
        If ParsePerform_(lbl, pTgt, pThru) Then
            PerformInto_ pTgt, pThru, node, stack, tr, out, secLabels
            Exit Sub
        End If
    End If

    ' GO TO <name>: forward exit-jump (Natural ESCAPE BOTTOM conversion
    ' style) - walk the target through its SECTION end (fall-through),
    ' then skip the rest of the current statement list. DEPENDING ON and
    ' unresolved targets stay generic actions.
    If v1 = "GO " Then
        If Left$(lbl, 6) = "GO TO " Then
            Dim gTgt As String
            gTgt = Mid$(lbl, 7)
            If InStr(gTgt, " ") = 0 Then
                ' GO TO a terminator section = abend, same as PERFORM
                If mTermSecs.Exists(gTgt) Then
                    tr.Add "Term", "abend:" & gTgt
                    tr.Add "TriggerLine", ln
                    AddEvent_ tr, "term", "ABEND-VIA " & gTgt, ln
                    AddCapped_ out, tr
                    Exit Sub
                End If
                Dim gOx As OrderedDict
                Set gOx = OwnerByName_(gTgt)
                If Not gOx Is Nothing Then
                    If Not mGotoSeen.Exists(gTgt) And Not OnStack_(stack, gTgt) Then
                        mGotoSeen.Add gTgt, True
                        AddEvent_ tr, "action", lbl, ln
                        AddEnterEvent_ tr, gTgt, secLabels, CLng(gOx.Item("line"))
                        Dim gSeed As Collection, gSubs As Collection, gSb As OrderedDict
                        Set gSeed = New Collection
                        gSeed.Add tr
                        stack.Add gTgt
                        Set gSubs = ApplyRange_(CLng(gOx.Item("line")), CLng(gOx.Item("secEnd")), stack, gSeed, secLabels)
                        stack.Remove stack.Count
                        For Each gSb In gSubs
                            gSb.Add "SkipRest", True   ' control was transferred
                            AddCapped_ out, gSb
                        Next gSb
                        Exit Sub
                    End If
                    ' resolved but already followed / cyclic: control still
                    ' transfers in reality - do not walk past the jump
                    AddEvent_ tr, "action", lbl, ln
                    tr.Add "SkipRest", True
                    AddCapped_ out, tr
                    Exit Sub
                End If
            End If
        End If
        AddEvent_ tr, "action", lbl, ln
        AddCapped_ out, tr
        Exit Sub
    End If

    ' READ <file> INTO <item>: the record lands in <item> - unknown value.
    ' A bare READ (record area only) is NOT tracked - the FD 01 stays a
    ' generic action and is never env-invalidated (accepted limitation).
    If v1 = "READ " Then
        Dim rdK As String, rdP As Long
        rdP = InStr(lbl, " INTO ")
        If rdP > 0 Then
            rdK = Trim$(Mid$(lbl, rdP + 6))
            rdP = InStr(rdK, " ")
            If rdP > 0 Then rdK = Left$(rdK, rdP - 1)
            rdK = UCase$(rdK)
        End If
        If Len(rdK) > 0 Then
            AddAssignEvent_ tr, rdK, lbl, "compute", ln
            Invalidate_ tr.Item("Env"), rdK
        Else
            AddEvent_ tr, "action", lbl, ln
        End If
        AddCapped_ out, tr
        Exit Sub
    End If

    ' arithmetic verbs change their targets - invalidate (and surface as
    ' computed assigns); without this a MOVE ZERO + ADD 1 counter keeps
    ' the stale constant and falsely prunes both arms of later tests
    If v1 = "ADD " Or v1 = "SUBTRACT " Or v1 = "MULTIPLY " Or v1 = "DIVIDE " Then
        Dim ats As Collection, av As Variant
        Set ats = ArithTargets_(lbl)
        For Each av In ats
            AddAssignEvent_ tr, CStr(av), lbl, "compute", ln
            Invalidate_ tr.Item("Env"), CStr(av)
        Next av
        If ats.Count = 0 Then AddEvent_ tr, "action", lbl, ln
        AddCapped_ out, tr
        Exit Sub
    End If

    ' CALL
    If v1 = "CALL " Then
    Set m = rxCall.Execute(lbl)
    If m.Count = 0 Then
        ' variable-name call: CALL WS-PGM-NAME USING ... (same submatch layout)
        Static rxCallV As Object
        If rxCallV Is Nothing Then
            Set rxCallV = CreateObject("VBScript.RegExp")
            rxCallV.Pattern = "^CALL\s+([A-Z0-9-]+)(\s+USING\s+(.+))?$"
            rxCallV.IgnoreCase = False
        End If
        Set m = rxCallV.Execute(lbl)
    End If
    If m.Count > 0 Then
        Dim tgt As String, params As String
        tgt = m.Item(0).SubMatches(0)
        params = m.Item(0).SubMatches(2)
        tr.Add "Calls", Cons_(tr.Item("Calls"), tgt)
        AddEvent_ tr, "call", lbl, ln
        ' pass 2 synthesized-failure walk: end the path right at this call
        If Len(mStopAtCall) > 0 Then
            If tgt = mStopAtCall Then
                AddEvent_ tr, "term", "CALL " & tgt & " -> ABNORMAL (synthesized)", ln
                tr.Add "Term", "synth:" & tgt
                tr.Add "TriggerLine", ln
                AddCapped_ out, tr
                Exit Sub
            End If
        End If
        ' remember the first call site (with the arm prefix reaching it) for
        ' synthesized failure cases - pass 2 replays the prefix to this call
        If Not mSynth.Exists(tgt) Then
            Dim sn As OrderedDict
            Set sn = New OrderedDict
            sn.Add "Line", ln
            sn.Add "Target", tgt
            sn.Add "ArmsAt", tr.Item("Arms")   ' cons head share - O(1)
            mSynth.Add tgt, sn
        End If
        If Len(params) > 0 Then
            Dim pa() As String, i As Long
            pa = Split(Trim$(params), " ")
            For i = LBound(pa) To UBound(pa)
                If Len(Trim$(pa(i))) > 0 Then Invalidate_ tr.Item("Env"), UCase$(Trim$(pa(i)))
            Next i
        End If
        AddCapped_ out, tr
        Exit Sub
    End If
    End If   ' v1 = "CALL "

    ' MOVE - supports multi-target form (MOVE v TO A B C). A quoted-literal
    ' source is matched first so literals containing " TO " keep working.
    Static rxMoveLit As Object
    If rxMoveLit Is Nothing Then
        Set rxMoveLit = CreateObject("VBScript.RegExp")
        rxMoveLit.Pattern = "^MOVE\s+('[^']*')\s+TO\s+(.+)$"
        rxMoveLit.IgnoreCase = False
    End If
    Dim srcTxt As String, dstList As String
    dstList = ""
    Set m = rxMoveLit.Execute(lbl)
    If m.Count > 0 Then
        srcTxt = m.Item(0).SubMatches(0)
        dstList = m.Item(0).SubMatches(1)
    Else
        Set m = rxMove.Execute(lbl)
        If m.Count > 0 Then
            srcTxt = m.Item(0).SubMatches(0)
            dstList = m.Item(0).SubMatches(1)
        End If
    End If
    If Len(dstList) > 0 Then
        Dim dts() As String, di As Long, dst As String, lit As String, anyDst As Boolean
        dts = Split(Trim$(dstList), " ")
        anyDst = False
        For di = LBound(dts) To UBound(dts)
            dst = Trim$(dts(di))
            If dst Like "[A-Z0-9]*" And InStr(dst, "(") = 0 Then
                anyDst = True
                AddAssignEvent_ tr, dst, srcTxt, "move", ln
                Invalidate_ tr.Item("Env"), dst
                If GetLiteral_(srcTxt, lit) Then
                    If dst <> mHavocItem Then tr.Item("Env").Add dst, lit
                End If
            End If
        Next di
        If Not anyDst Then AddEvent_ tr, "action", lbl, ln
        AddCapped_ out, tr
        Exit Sub
    End If

    ' COMPUTE / STRING / INITIALIZE / ACCEPT - assignment with unknown value
    Set m = rxComp.Execute(lbl)
    If m.Count > 0 Then
        AddAssignEvent_ tr, m.Item(0).SubMatches(0), m.Item(0).SubMatches(1), "compute", ln
        Invalidate_ tr.Item("Env"), m.Item(0).SubMatches(0)
        AddCapped_ out, tr
        Exit Sub
    End If
    Set m = rxStr.Execute(lbl)
    If m.Count > 0 Then
        AddAssignEvent_ tr, m.Item(0).SubMatches(1), "STRING " & m.Item(0).SubMatches(0), "string", ln
        Invalidate_ tr.Item("Env"), m.Item(0).SubMatches(1)
        AddCapped_ out, tr
        Exit Sub
    End If
    Set m = rxInit.Execute(lbl)
    If m.Count > 0 Then
        AddAssignEvent_ tr, m.Item(0).SubMatches(0), "INITIALIZE", "init", ln
        Invalidate_ tr.Item("Env"), m.Item(0).SubMatches(0)
        AddCapped_ out, tr
        Exit Sub
    End If
    Set m = rxAcc.Execute(lbl)
    If m.Count > 0 Then
        AddAssignEvent_ tr, m.Item(0).SubMatches(0), "ACCEPT", "accept", ln
        Invalidate_ tr.Item("Env"), m.Item(0).SubMatches(0)
        AddCapped_ out, tr
        Exit Sub
    End If

    ' EXEC and anything else: generic event (EXEC is normally handled by the
    ' fast-dispatch Case Else above; this is the fallback for a listed verb that
    ' fell through its handler)
    If Left$(lbl, 5) = "EXEC " Then
        AddEvent_ tr, "exec", lbl, ln
    Else
        AddEvent_ tr, "action", lbl, ln
    End If
    AddCapped_ out, tr
End Sub

Private Sub PerformInto_(ByVal fromName As String, ByVal thruName As String, ByVal node As OrderedDict, _
                         ByVal stack As Collection, ByVal tr As OrderedDict, ByVal out As Collection, _
                         ByVal secLabels As OrderedDict)
    Dim ln As Long
    ln = CLng(node.Item("startLine"))

    If mTermSecs.Exists(fromName) Then
        tr.Add "Term", "abend:" & fromName
        tr.Add "TriggerLine", ln
        AddEvent_ tr, "term", "ABEND-VIA " & fromName, ln
        AddCapped_ out, tr
        Exit Sub
    End If

    Dim ox As OrderedDict
    Set ox = OwnerByName_(fromName)
    If ox Is Nothing Or OnStack_(stack, fromName) Then
        AddEvent_ tr, "action", "PERFORM " & fromName, ln
        AddCapped_ out, tr
        Exit Sub
    End If

    Dim hi As Long
    If CStr(ox.Item("kind")) = "section" Then hi = CLng(ox.Item("secEnd")) Else hi = CLng(ox.Item("ownerEnd"))
    If Len(thruName) > 0 Then
        Dim oy As OrderedDict
        Set oy = OwnerByName_(thruName)
        If Not oy Is Nothing Then
            If CStr(oy.Item("kind")) = "section" Then hi = CLng(oy.Item("secEnd")) Else hi = CLng(oy.Item("ownerEnd"))
        End If
    End If

    AddEnterEvent_ tr, fromName, secLabels, CLng(ox.Item("line"))
    stack.Add fromName
    Dim seed As Collection, subs As Collection, sb As OrderedDict
    Set seed = New Collection
    seed.Add tr
    Set subs = ApplyRange_(CLng(ox.Item("line")), hi, stack, seed, secLabels)
    stack.Remove stack.Count
    For Each sb In subs
        ' a GO TO inside the callee ended at its section exit - control
        ' returns to this caller, so the skip flag is cleared here
        If TrSkips_(sb) Then sb.Add "SkipRest", False
        AddCapped_ out, sb
    Next sb
End Sub

Private Function TrSkips_(ByVal tr As OrderedDict) As Boolean
    TrSkips_ = False
    If tr.Exists("SkipRest") Then TrSkips_ = CBool(tr.Item("SkipRest"))
End Function

' "<section>|<callers>|<refs>" for an unreached arm: the owning section,
' who the static call graph says invokes it (the section OR any paragraph
' inside it), and - when the graph has nothing - raw source lines that
' mention the section name (the invocation form verbatim: GO TO ...
' DEPENDING, SORT INPUT PROCEDURE, etc.), so the matrix itself shows HOW
' the region is entered without anyone grepping the source by hand.
Private Function NoCtxDetail_(ByVal lineNo As Long, ByVal callG As OrderedDict, _
                              ByVal lines As Collection) As String
    Dim secName As String, o As OrderedDict, secLo As Long, secHi As Long
    secName = ""
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" Then
            If lineNo >= CLng(o.Item("line")) And lineNo <= CLng(o.Item("secEnd")) Then
                secName = CStr(o.Item("name"))
                secLo = CLng(o.Item("line"))
                secHi = CLng(o.Item("secEnd"))
                Exit For
            End If
        End If
    Next o
    If Len(secName) = 0 Then
        NoCtxDetail_ = "||"
        Exit Function
    End If
    ' every label belonging to the section counts as an entry point
    Dim names As String
    names = "," & secName & ","
    For Each o In mOwners
        If CStr(o.Item("kind")) = "para" Then
            If CLng(o.Item("line")) >= secLo And CLng(o.Item("line")) <= secHi Then
                names = names & CStr(o.Item("name")) & ","
            End If
        End If
    Next o
    Dim callers As String, e As OrderedDict, nHit As Long
    callers = ""
    If Not callG Is Nothing Then
        If callG.Exists("edges") Then
            For Each e In callG.Item("edges")
                ' self-references from inside the dead region do not count
                If CLng(e.Item("line")) >= secLo And CLng(e.Item("line")) <= secHi Then
                    ' skip
                ElseIf InStr(names, "," & UCase$(CStr(e.Item("to"))) & ",") > 0 And _
                   Left$(CStr(e.Item("kind")), 7) = "perform" Then
                    If InStr("," & callers & ",", "," & CStr(e.Item("from")) & ",") = 0 Then
                        nHit = nHit + 1
                        If nHit <= 3 Then
                            If Len(callers) > 0 Then callers = callers & ","
                            callers = callers & CStr(e.Item("from"))
                        End If
                    End If
                End If
            Next e
        End If
    End If
    ' no graph edge: quote the source lines that reference the name
    Dim refs As String, le As OrderedDict, txt As String, p As Long, ch As String
    Dim nRef As Long, hit As Boolean, q As String
    refs = ""
    If Len(callers) = 0 Then
        For Each le In lines
            If CLng(le.Item("Number")) < secLo Or CLng(le.Item("Number")) > secHi Then
                txt = UCase$(CStr(le.Item("Text")))
                ' token-boundary match on both sides, any occurrence
                hit = False
                p = InStr(txt, secName)
                Do While p > 0 And Not hit
                    ch = Mid$(txt & " ", p + Len(secName), 1)
                    If ch Like "[!A-Z0-9-]" Then
                        If p = 1 Then
                            hit = True
                        ElseIf Mid$(txt, p - 1, 1) Like "[!A-Z0-9-]" Then
                            hit = True
                        End If
                    End If
                    If Not hit Then p = InStr(p + 1, txt, secName)
                Loop
                If hit Then
                    q = Replace(Left$(CStr(le.Item("Text")), 45), "|", "/")
                    q = Replace(q, "~", "-")
                    If Len(refs) > 0 Then refs = refs & "~"
                    refs = refs & "L" & CLng(le.Item("Number")) & ":" & q
                    nRef = nRef + 1
                    If nRef >= 2 Then Exit For
                End If
            End If
        Next le
    End If
    NoCtxDetail_ = secName & "|" & callers & "|" & refs
End Function

Private Function OnStack_(ByVal stack As Collection, ByVal nm As String) As Boolean
    Dim v As Variant
    OnStack_ = False
    For Each v In stack
        If CStr(v) = nm Then OnStack_ = True
    Next v
End Function

Private Sub ApplyEvaluate_(ByVal node As OrderedDict, ByVal stack As Collection, ByVal tr As OrderedDict, _
                           ByVal out As Collection, ByVal secLabels As OrderedDict)
    Dim expr As String, known As String, hasKnown As Boolean
    expr = CStr(node.Item("expression"))
    hasKnown = False
    If EnvHas_(tr.Item("Env"), expr) Then
        known = NormVal_(CStr(tr.Item("Env").Item(expr)))
        hasKnown = True
    End If

    Dim cs As Collection, w As OrderedDict, lit As String
    Set cs = node.Item("cases")
    Dim litAll As Boolean
    litAll = True
    Dim wi As Long
    ' if every WHEN value is a literal and expr known, select just one arm
    If hasKnown Then
        For wi = 1 To cs.Count
            Set w = cs(wi)
            If CStr(w.Item("condition")) <> "OTHER" Then
                If Not GetLiteral_(CStr(w.Item("condition")), lit) Then litAll = False
            End If
        Next wi
    Else
        litAll = False
    End If

    Dim matchedIdx As Long
    matchedIdx = 0
    If litAll Then
        For wi = 1 To cs.Count
            Set w = cs(wi)
            If CStr(w.Item("condition")) <> "OTHER" Then
                If GetLiteral_(CStr(w.Item("condition")), lit) Then
                    If NormVal_(lit) = known Then
                        matchedIdx = wi
                        Exit For
                    End If
                End If
            End If
        Next wi
    End If

    Dim hasOther As Boolean
    hasOther = False
    For wi = 1 To cs.Count
        If CStr(cs(wi).Item("condition")) = "OTHER" Then hasOther = True
    Next wi

    ' directed: pick exactly one arm. "allowed" reflects the env pruning the
    ' forking version applied (known literal subject narrows the choices).
    Dim needIdx As Long, needSkip As Boolean
    needIdx = 0
    needSkip = mNeed.Exists(CStr(node.Item("id")) & ":skip")
    For wi = 1 To cs.Count
        If mNeed.Exists(CStr(cs(wi).Item("id"))) Then needIdx = wi
    Next wi

    Dim selIdx As Long, selSkip As Boolean
    selIdx = 0
    selSkip = False
    ' tier 1: pass-2 replay
    If Not mReplayList Is Nothing Then
        If mReplayIdx <= mReplayList.Count Then
            Dim rt As String
            rt = CStr(mReplayList(mReplayIdx))
            For wi = 1 To cs.Count
                If rt = CStr(cs(wi).Item("id")) Then
                    selIdx = wi
                    mReplayIdx = mReplayIdx + 1
                    Exit For
                End If
            Next wi
            If selIdx = 0 Then
                If rt = CStr(node.Item("id")) & ":skip" Then
                    selSkip = True
                    mReplayIdx = mReplayIdx + 1
                End If
            End If
        End If
    End If
    If selIdx = 0 And Not selSkip Then
        If needIdx > 0 Then
            If WhenAllowed_(needIdx, cs, litAll, matchedIdx) Or mForce Then
                selIdx = needIdx
            Else
                mMissed = True
                mMissCond = expr & " = " & CStr(cs(needIdx).Item("condition"))
                mMissTok = CStr(cs(needIdx).Item("id"))
            End If
        ElseIf needSkip Then
            If hasOther Or (litAll And matchedIdx > 0) Then
                mMissed = True
                mMissCond = expr & " = (no match)"
            Else
                selSkip = True
            End If
        ElseIf mSweep Then
            ' tier 3: uncovered non-abend WHEN first, then the heaviest
            ' subtree (abend WHENs stay for the targeted fallback walks)
            For wi = 1 To cs.Count
                If mSweepUncov.Exists(CStr(cs(wi).Item("id"))) Then
                    If WhenAllowed_(wi, cs, litAll, matchedIdx) Then
                        If Not BlockAbends_(cs(wi).Item("children")) Then
                            selIdx = wi
                            Exit For
                        End If
                    End If
                End If
            Next wi
            If selIdx = 0 And Not hasOther Then
                If mSweepUncov.Exists(CStr(node.Item("id")) & ":skip") And Not (litAll And matchedIdx > 0) Then selSkip = True
            End If
            If selIdx = 0 And Not selSkip Then
                Dim bw As Long, ww As Long
                bw = 0
                For wi = 1 To cs.Count
                    ww = 0
                    If mSweepW.Exists(CStr(cs(wi).Item("id"))) Then ww = CLng(mSweepW.Item(CStr(cs(wi).Item("id"))))
                    If ww > bw Then
                        If WhenAllowed_(wi, cs, litAll, matchedIdx) Then
                            bw = ww
                            selIdx = wi
                        End If
                    End If
                Next wi
            End If
        End If
    End If
    If selIdx = 0 And Not selSkip Then
        ' default choice
        If litAll Then
            If matchedIdx > 0 Then
                selIdx = matchedIdx
            ElseIf hasOther Then
                For wi = 1 To cs.Count
                    If CStr(cs(wi).Item("condition")) = "OTHER" Then selIdx = wi
                Next wi
            Else
                selSkip = True
            End If
        Else
            If mPrefElse Then
                If hasOther Then
                    For wi = 1 To cs.Count
                        If CStr(cs(wi).Item("condition")) = "OTHER" Then selIdx = wi
                    Next wi
                ElseIf cs.Count > 0 Then
                    selIdx = cs.Count
                Else
                    selSkip = True
                End If
            Else
                If cs.Count > 0 Then selIdx = 1 Else selSkip = True
            End If
        End If
    End If

    Dim seed As Collection, subs As Collection, sb As OrderedDict
    If selIdx > 0 Then
        Set w = cs(selIdx)
        AddArmEvent_ tr, CStr(w.Item("id")), "WHEN", expr & " = " & CStr(w.Item("condition")), CLng(w.Item("startLine"))
        Set seed = New Collection
        seed.Add tr
        Set subs = ApplyNodes_(w.Item("children"), stack, seed, secLabels)
        For Each sb In subs: AddCapped_ out, sb: Next sb
    ElseIf selSkip And Not hasOther Then
        AddArmEvent_ tr, CStr(node.Item("id")) & ":skip", "SKIP", expr & " = (no match)", CLng(node.Item("startLine"))
        AddCapped_ out, tr
    Else
        mMissed = True
    End If
End Sub

Private Function WhenAllowed_(ByVal wi As Long, ByVal cs As Collection, _
                              ByVal litAll As Boolean, ByVal matchedIdx As Long) As Boolean
    If Not litAll Then
        WhenAllowed_ = True
    ElseIf matchedIdx > 0 Then
        WhenAllowed_ = (wi = matchedIdx)
    Else
        WhenAllowed_ = (CStr(cs(wi).Item("condition")) = "OTHER")
    End If
End Function

Private Sub ApplySearch_(ByVal node As OrderedDict, ByVal stack As Collection, ByVal tr As OrderedDict, _
                         ByVal out As Collection, ByVal secLabels As OrderedDict)
    ' directed: exactly one arm - a required WHEN / AT END, else default
    ' (default = first WHEN i.e. "found"; ELSE-pref seeds take AT END/skip)
    Dim cs As Collection, wi As Long, w As OrderedDict
    Set cs = node.Item("cases")
    Dim hasAtEnd As Boolean
    hasAtEnd = Not IsNull(node.Item("atEndLine"))

    Dim needIdx As Long, needAtEnd As Boolean, needSkip As Boolean
    needIdx = 0
    needAtEnd = mNeed.Exists(CStr(node.Item("id")) & ":atend")
    needSkip = mNeed.Exists(CStr(node.Item("id")) & ":skip")
    For wi = 1 To cs.Count
        If mNeed.Exists(CStr(cs(wi).Item("id"))) Then needIdx = wi
    Next wi

    Dim selIdx As Long, selEnd As Boolean
    selIdx = 0
    selEnd = False
    Dim chosen As Boolean
    chosen = False
    ' tier 1: pass-2 replay
    If Not mReplayList Is Nothing Then
        If mReplayIdx <= mReplayList.Count Then
            Dim rt As String
            rt = CStr(mReplayList(mReplayIdx))
            For wi = 1 To cs.Count
                If rt = CStr(cs(wi).Item("id")) Then
                    selIdx = wi
                    chosen = True
                    mReplayIdx = mReplayIdx + 1
                    Exit For
                End If
            Next wi
            If Not chosen Then
                If rt = CStr(node.Item("id")) & ":atend" Or rt = CStr(node.Item("id")) & ":skip" Then
                    selEnd = True
                    chosen = True
                    mReplayIdx = mReplayIdx + 1
                End If
            End If
        End If
    End If
    If Not chosen Then
        If needIdx > 0 Then
            selIdx = needIdx
        ElseIf needAtEnd Or needSkip Then
            selEnd = True
        ElseIf mSweep Then
            ' tier 3: uncovered non-abend WHEN / AT END first, then the
            ' heaviest subtree (abend arms stay for the fallback walks)
            For wi = 1 To cs.Count
                If mSweepUncov.Exists(CStr(cs(wi).Item("id"))) Then
                    If Not BlockAbends_(cs(wi).Item("children")) Then
                        selIdx = wi
                        Exit For
                    End If
                End If
            Next wi
            If selIdx = 0 Then
                If mSweepUncov.Exists(CStr(node.Item("id")) & ":atend") Or _
                   mSweepUncov.Exists(CStr(node.Item("id")) & ":skip") Then
                    selEnd = True
                End If
            End If
            If selIdx = 0 And Not selEnd Then
                Dim bw As Long, ww As Long
                bw = 0
                For wi = 1 To cs.Count
                    ww = 0
                    If mSweepW.Exists(CStr(cs(wi).Item("id"))) Then ww = CLng(mSweepW.Item(CStr(cs(wi).Item("id"))))
                    If ww > bw Then
                        bw = ww
                        selIdx = wi
                    End If
                Next wi
                If selIdx = 0 Then
                    If cs.Count > 0 Then selIdx = 1 Else selEnd = True
                End If
            End If
        ElseIf mPrefElse Or cs.Count = 0 Then
            selEnd = True
        Else
            selIdx = 1
        End If
    End If

    Dim seed As Collection, subs As Collection, sb As OrderedDict
    If selIdx > 0 Then
        Set w = cs(selIdx)
        AddArmEvent_ tr, CStr(w.Item("id")), "WHEN", CStr(w.Item("condition")), CLng(w.Item("startLine"))
        Set seed = New Collection
        seed.Add tr
        Set subs = ApplyNodes_(w.Item("children"), stack, seed, secLabels)
        For Each sb In subs: AddCapped_ out, sb: Next sb
    ElseIf hasAtEnd Then
        AddArmEvent_ tr, CStr(node.Item("id")) & ":atend", "AT END", "SEARCH " & CStr(node.Item("tableExpr")), CLng(node.Item("startLine"))
        Set seed = New Collection
        seed.Add tr
        Set subs = ApplyNodes_(node.Item("atEndChildren"), stack, seed, secLabels)
        For Each sb In subs: AddCapped_ out, sb: Next sb
    Else
        AddArmEvent_ tr, CStr(node.Item("id")) & ":skip", "AT END(skip)", "SEARCH " & CStr(node.Item("tableExpr")), CLng(node.Item("startLine"))
        AddCapped_ out, tr
    End If
End Sub

'======================================================================
' Arms catalog
'======================================================================
Private Sub CollectArms_(ByVal nodes As Collection, ByVal arms As Collection)
    Dim n As OrderedDict, t As String
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "if" Then
            AddArm_ arms, CStr(n.Item("id")) & ":then", CLng(n.Item("startLine")), "IF " & CStr(n.Item("condition")) & " [THEN]", _
                    FirstChildLine_(n.Item("thenChildren"), CLng(n.Item("startLine")))
            AddArm_ arms, CStr(n.Item("id")) & ":else", CLng(n.Item("startLine")), "IF " & CStr(n.Item("condition")) & " [ELSE]", _
                    FirstChildLine_(n.Item("elseChildren"), CLng(n.Item("startLine")))
            CollectArms_ n.Item("thenChildren"), arms
            CollectArms_ n.Item("elseChildren"), arms
        ElseIf t = "evaluate" Then
            Dim cs As Collection, wi As Long, w As OrderedDict, hasOther As Boolean
            Set cs = n.Item("cases")
            hasOther = False
            For wi = 1 To cs.Count
                Set w = cs(wi)
                AddArm_ arms, CStr(w.Item("id")), CLng(w.Item("startLine")), "EVALUATE " & CStr(n.Item("expression")) & " WHEN " & CStr(w.Item("condition")), _
                        CLng(w.Item("startLine"))
                If CStr(w.Item("condition")) = "OTHER" Then hasOther = True
                CollectArms_ w.Item("children"), arms
            Next wi
            If Not hasOther Then
                AddArm_ arms, CStr(n.Item("id")) & ":skip", CLng(n.Item("startLine")), "EVALUATE " & CStr(n.Item("expression")) & " [no-match skip]", _
                        CLng(n.Item("startLine"))
            End If
        ElseIf t = "search" Then
            If Not IsNull(n.Item("atEndLine")) Then
                AddArm_ arms, CStr(n.Item("id")) & ":atend", CLng(n.Item("startLine")), "SEARCH " & CStr(n.Item("tableExpr")) & " [AT END]", _
                        FirstChildLine_(n.Item("atEndChildren"), CLng(n.Item("startLine")))
            Else
                AddArm_ arms, CStr(n.Item("id")) & ":skip", CLng(n.Item("startLine")), "SEARCH " & CStr(n.Item("tableExpr")) & " [AT END skip]", _
                        CLng(n.Item("startLine"))
            End If
            Dim sc As Collection, si As Long, sw As OrderedDict
            Set sc = n.Item("cases")
            For si = 1 To sc.Count
                Set sw = sc(si)
                AddArm_ arms, CStr(sw.Item("id")), CLng(sw.Item("startLine")), "SEARCH WHEN " & CStr(sw.Item("condition")), _
                        CLng(sw.Item("startLine"))
                CollectArms_ sw.Item("children"), arms
            Next si
        End If
    Next n
End Sub

' the line where a tester would mark this arm in a source listing: the
' first statement INSIDE the arm (falls back to the branch line itself
' for empty arms / skip pseudo-arms)
Private Function FirstChildLine_(ByVal children As Collection, ByVal fallback As Long) As Long
    FirstChildLine_ = fallback
    If children Is Nothing Then Exit Function
    If children.Count > 0 Then FirstChildLine_ = CLng(children(1).Item("startLine"))
End Function

Private Sub AddArm_(ByVal arms As Collection, ByVal token As String, ByVal lineNo As Long, _
                    ByVal disp As String, ByVal markLine As Long)
    Dim a As OrderedDict
    Set a = New OrderedDict
    a.Add "Token", token
    a.Add "Line", lineNo
    a.Add "Disp", disp
    a.Add "MarkLine", markLine
    arms.Add a
End Sub

'======================================================================
' Case selection (C1 greedy + abend grouping + synthesized failures)
'======================================================================
Private Function SelectCases_(ByVal normals As Collection, ByVal abends As Collection, _
                              ByVal arms As Collection) As Collection
    Dim cases As Collection
    Set cases = New Collection

    ' every candidate already carries its materialized ArmsL (AddCand_)
    Dim coverable As OrderedDict, tr As OrderedDict, v As Variant

    ' C1 greedy over normal traces
    Set coverable = New OrderedDict
    For Each tr In normals
        For Each v In tr.Item("ArmsL")
            If Not coverable.Exists(CStr(v)) Then coverable.Add CStr(v), True
        Next v
    Next tr

    Dim uncov As OrderedDict, ks As Collection
    Set uncov = New OrderedDict
    Set ks = coverable.Keys
    For Each v In ks
        uncov.Add CStr(v), True
    Next v

    Dim picked As Collection
    Set picked = New Collection
    Do While CountLive_(uncov) > 0
        Dim best As OrderedDict, bestGain As Long, gain As Long
        Set best = Nothing
        bestGain = 0
        For Each tr In normals
            gain = 0
            For Each v In tr.Item("ArmsL")
                If uncov.Exists(CStr(v)) Then
                    If Not IsEmpty(uncov.Item(CStr(v))) Then gain = gain + 1
                End If
            Next v
            If gain > bestGain Then
                Set best = tr
                bestGain = gain
            End If
        Next tr
        If best Is Nothing Then Exit Do
        picked.Add best
        For Each v In best.Item("ArmsL")
            If uncov.Exists(CStr(v)) Then uncov.Add CStr(v), Empty
        Next v
    Loop
    ' straight-line program (no branch arms on any normal path): still emit
    ' one normal case so the scenario list is never silently empty
    If picked.Count = 0 And normals.Count > 0 Then picked.Add normals(1)

    ' code-derived abend cases: one per (final arm, terminator), shortest.
    ' the cons head IS the last arm, so this needs no materialization.
    Dim groups As OrderedDict, key As String
    Set groups = New OrderedDict
    For Each tr In abends
        Dim ah As ConsList
        Set ah = tr.Item("Arms")
        If Not ah Is Nothing Then
            key = CStr(ah.V) & "|" & CStr(tr.Item("Term"))
        Else
            key = "(none)|" & CStr(tr.Item("Term"))
        End If
        If Not groups.Exists(key) Then
            groups.Add key, tr
        ElseIf ConsCount_(ah) < ConsCount_(groups.Item(key).Item("Arms")) Then
            groups.Add key, tr
        End If
    Next tr

    ' synthesized call-failure specs (re-walked with stop-at-call in pass 2)
    Dim synths As Collection
    Set synths = New Collection
    Dim sk As Collection, sv As Variant, sn As OrderedDict
    Set sk = mSynth.Keys
    For Each sv In sk
        Set sn = mSynth.Item(CStr(sv))
        Dim ss As OrderedDict
        Set ss = New OrderedDict
        ss.Add "Kind", "synth"
        ss.Add "NeedList", ConsToList_(sn.Item("ArmsAt"))
        ss.Add "SynthTarget", CStr(sn.Item("Target"))
        ss.Add "TriggerLine", CLng(sn.Item("Line"))
        synths.Add ss
    Next sv

    ' abnormal = code-abend specs + synth specs, ordered by trigger line
    Dim abnormal As Collection
    Set abnormal = New Collection
    Dim gk As Collection, gv As Variant
    Set gk = groups.Keys
    For Each gv In gk
        Set tr = groups.Item(CStr(gv))
        Dim ab As OrderedDict
        Set ab = New OrderedDict
        ab.Add "Kind", "abend"
        ab.Add "NeedList", tr.Item("ArmsL")
        If tr.Exists("EntrySec") Then ab.Add "EntrySec", CStr(tr.Item("EntrySec"))
        ab.Add "SynthTarget", ""
        ab.Add "TriggerLine", CLng(tr.Item("TriggerLine"))
        abnormal.Add ab
    Next gv
    Dim sy As OrderedDict
    For Each sy In synths
        abnormal.Add sy
    Next sy
    Set abnormal = SortByTrigger_(abnormal)

    ' assemble pass-2 specs (cases are re-walked with events recorded)
    Dim serial As Long, normSerial As Long, abSerial As Long
    serial = 0
    Dim sp As OrderedDict
    For Each tr In picked
        serial = serial + 1
        normSerial = normSerial + 1
        Set sp = New OrderedDict
        sp.Add "Kind", "normal"
        sp.Add "Id", "TC" & serial
        sp.Add "KindSerial", normSerial
        sp.Add "NeedList", tr.Item("ArmsL")
        If tr.Exists("EntrySec") Then sp.Add "EntrySec", CStr(tr.Item("EntrySec"))
        sp.Add "SynthTarget", ""
        cases.Add sp
    Next tr
    For Each sp In abnormal
        serial = serial + 1
        abSerial = abSerial + 1
        sp.Add "Id", "TC" & serial          ' OrderedDict.Add updates if present
        sp.Add "KindSerial", abSerial
        cases.Add sp
    Next sp

    Set SelectCases_ = cases
End Function

Private Function CountLive_(ByVal d As OrderedDict) As Long
    Dim ks As Collection, v As Variant, n As Long
    Set ks = d.Keys
    For Each v In ks
        If Not IsEmpty(d.Item(CStr(v))) Then n = n + 1
    Next v
    CountLive_ = n
End Function

Private Function SortByTrigger_(ByVal c As Collection) As Collection
    Dim res As Collection
    Set res = New Collection
    Dim used() As Boolean, n As Long, k As Long, i As Long, best As Long, bestLine As Long
    n = c.Count
    Set SortByTrigger_ = res
    If n = 0 Then Exit Function
    ReDim used(1 To n)
    For k = 1 To n
        best = 0: bestLine = 2000000000
        For i = 1 To n
            If Not used(i) Then
                If CLng(c(i).Item("TriggerLine")) < bestLine Then
                    bestLine = CLng(c(i).Item("TriggerLine")): best = i
                End If
            End If
        Next i
        used(best) = True
        res.Add c(best)
    Next k
    Set SortByTrigger_ = res
End Function

Private Function BuildCase_(ByVal tr As OrderedDict, ByVal id As String, ByVal kind As String, _
                            ByVal kindSerial As Long) As OrderedDict
    Dim c As OrderedDict
    Set c = New OrderedDict
    c.Add "id", id
    c.Add "kind", kind            ' normal / abend / synth
    c.Add "kindSerial", kindSerial
    c.Add "term", tr.Item("Term")
    c.Add "triggerLine", tr.Item("TriggerLine")
    ' materialize the cons lists for this selected case only
    c.Add "arms", ConsToList_(tr.Item("Arms"))
    Dim evs As Collection
    Set evs = ConsToList_(tr.Item("Events"))
    c.Add "events", evs
    ' final action = last lined non-term event (the user's definition: the
    ' last statement executed right before the program really ends)
    Dim i As Long, fl As Long, e As OrderedDict
    fl = 0
    For i = evs.Count To 1 Step -1
        Set e = evs(i)
        If CStr(e.Item("Kind")) <> "term" And CStr(e.Item("Kind")) <> "enter" And CStr(e.Item("Kind")) <> "arm" Then
            fl = CLng(e.Item("Line"))
            Exit For
        End If
    Next i
    c.Add "finalLine", fl
    ' the actual terminator verb/text (GOBACK vs STOP RUN etc.) for labels
    Dim tv As String
    tv = ""
    For i = evs.Count To 1 Step -1
        Set e = evs(i)
        If CStr(e.Item("Kind")) = "term" Then
            tv = CStr(e.Item("Text"))
            Exit For
        End If
    Next i
    c.Add "termVerb", tv
    Set BuildCase_ = c
End Function
