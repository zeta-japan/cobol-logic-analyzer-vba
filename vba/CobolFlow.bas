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
'   normalPaths / abendPaths - Long counts (after pruning)
'   truncated - Boolean (MAX_TRACES cap hit; coverage may be partial)
'   entryName - String

Option Explicit

Private Const MAX_TRACES As Long = 200
Private mOps As Long   ' heartbeat counter: keep Excel responsive on big programs

Private mNodes As Collection      ' AST root nodes
Private mOwners As Collection     ' {name,line,kind,ownerEnd,secEnd} sorted
Private mCut As OrderedDict       ' plain-PERFORM target names
Private mTermSecs As OrderedDict  ' registered terminator section names
Private mDesc As OrderedDict      ' item -> Collection of descendant names
Private mAnc As OrderedDict       ' item -> Collection of ancestor names
Private mCondItems As OrderedDict ' identifiers used in any branch condition
Private mSynth As OrderedDict     ' call target -> {Trace, Line} first site
Private mTruncated As Boolean

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
    Dim ow As OrderedDict
    For Each ow In mOwners
        If CStr(ow.Item("kind")) = "section" And InStr(CStr(ow.Item("name")), "ABEND") > 0 Then
            If Not mTermSecs.Exists(CStr(ow.Item("name"))) Then
                mTermSecs.Add CStr(ow.Item("name")), True
                termsApplied.Add MakeTermInfo_(CStr(ow.Item("name")), "auto")
            End If
        End If
    Next ow

    BuildCut_
    BuildGroupMaps_ lines
    BuildCondItems_ mNodes
    Set mSynth = New OrderedDict
    mTruncated = False
    mOps = 0

    ' stage labels: comment with full-width parens just above a section header
    Dim secLabels As OrderedDict
    Set secLabels = BuildSecLabels_(norm)

    ' enumerate from the entry section
    Dim entry As OrderedDict
    Set entry = FindEntry_()
    Dim allTraces As Collection
    Set allTraces = New Collection
    If Not entry Is Nothing And mNodes.Count > 0 Then
        Dim t0 As OrderedDict
        Set t0 = NewTrace_()
        AddEnterEvent_ t0, CStr(entry.Item("name")), secLabels, CLng(entry.Item("line"))
        Dim seed As Collection, stack As Collection
        Set seed = New Collection
        seed.Add t0
        Set stack = New Collection
        stack.Add CStr(entry.Item("name"))
        Set allTraces = ApplyRange_(CLng(entry.Item("line")), CLng(entry.Item("secEnd")), stack, seed, secLabels)
    End If

    ' split normal / abend
    Dim normals As Collection, abends As Collection, tr As OrderedDict
    Set normals = New Collection
    Set abends = New Collection
    For Each tr In allTraces
        If CStr(tr.Item("Term")) = "goback" Then
            normals.Add tr
        ElseIf Left$(CStr(tr.Item("Term")), 6) = "abend:" Then
            abends.Add tr
        End If
    Next tr

    Dim arms As Collection
    Set arms = New Collection
    CollectArms_ mNodes, arms

    Dim cases As Collection
    Set cases = SelectCases_(normals, abends, arms)

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "cases", cases
    result.Add "arms", arms
    result.Add "normalPaths", normals.Count
    result.Add "abendPaths", abends.Count
    result.Add "truncated", mTruncated
    result.Add "descMap", mDesc   ' item -> descendants (for downstream IO derivation)
    result.Add "termsApplied", termsApplied
    If entry Is Nothing Then
        result.Add "entryName", ""
    Else
        result.Add "entryName", CStr(entry.Item("name"))
    End If
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

Private Sub BuildCut_()
    Set mCut = New OrderedDict
    CollectPerformTargets_ mNodes
End Sub

Private Sub CollectPerformTargets_(ByVal nodes As Collection)
    Static rxP As Object
    If rxP Is Nothing Then
        Set rxP = CreateObject("VBScript.RegExp")
        rxP.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9-]*)$"
        rxP.IgnoreCase = False
    End If
    Dim n As OrderedDict, t As String, m As Object
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "action" Then
            Set m = rxP.Execute(CStr(n.Item("label")))
            If m.Count > 0 Then
                If Not mCut.Exists(m.Item(0).SubMatches(0)) Then mCut.Add m.Item(0).SubMatches(0), True
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

Private Function OwnerByName_(ByVal nm As String) As OrderedDict
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

Private Function CloneTrace_(ByVal t As OrderedDict) As OrderedDict
    Dim c As OrderedDict
    Set c = New OrderedDict
    c.Add "Arms", t.Item("Arms")     ' cons heads are immutable - share them
    c.Add "Events", t.Item("Events")
    c.Add "Calls", t.Item("Calls")
    Dim env As OrderedDict, src As OrderedDict, ks As Collection, v As Variant
    Set env = New OrderedDict
    Set src = t.Item("Env")
    Set ks = src.Keys
    For Each v In ks
        env.Add CStr(v), src.Item(CStr(v))
    Next v
    c.Add "Env", env
    c.Add "Term", t.Item("Term")
    c.Add "TriggerLine", t.Item("TriggerLine")
    Set CloneTrace_ = c
End Function

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
    Set Cons_ = n
End Function

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
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "Kind", kind
    e.Add "Text", text
    e.Add "Line", lineNo
    t.Add "Events", Cons_(t.Item("Events"), e)
End Sub

Private Sub AddEnterEvent_(ByVal t As OrderedDict, ByVal secName As String, ByVal secLabels As OrderedDict, ByVal lineNo As Long)
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
    t.Add "Arms", Cons_(t.Item("Arms"), token)
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
    Set ApplyRange_ = ApplyNodes_(NodesInRange_(lo, CapHi_(lo, hi)), stack, traces, secLabels)
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
    If (mOps And 2047) = 0 Then DoEvents   ' stay responsive on big programs
    Dim out As Collection
    Set out = New Collection
    Dim tr As OrderedDict, t As String
    t = CStr(node.Item("type"))

    For Each tr In traces
        If CStr(tr.Item("Term")) <> "" Then
            AddCapped_ out, tr
        ElseIf t = "action" Then
            ApplyAction_ node, stack, tr, out, secLabels
        ElseIf t = "if" Then
            Dim okT As Boolean, okE As Boolean
            TestFeasible_ CStr(node.Item("condition")), tr.Item("Env"), okT, okE
            Dim c1 As OrderedDict, subs As Collection, sb As OrderedDict, seed As Collection
            If okT Then
                Set c1 = CloneTrace_(tr)
                AddArmEvent_ c1, CStr(node.Item("id")) & ":then", "THEN", CStr(node.Item("condition")), CLng(node.Item("startLine"))
                Set seed = New Collection
                seed.Add c1
                Set subs = ApplyNodes_(node.Item("thenChildren"), stack, seed, secLabels)
                For Each sb In subs: AddCapped_ out, sb: Next sb
            End If
            If okE Then
                Set c1 = CloneTrace_(tr)
                AddArmEvent_ c1, CStr(node.Item("id")) & ":else", "ELSE", CStr(node.Item("condition")), CLng(node.Item("startLine"))
                Set seed = New Collection
                seed.Add c1
                Set subs = ApplyNodes_(node.Item("elseChildren"), stack, seed, secLabels)
                For Each sb In subs: AddCapped_ out, sb: Next sb
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

Private Sub AddCapped_(ByVal out As Collection, ByVal tr As OrderedDict)
    If out.Count >= MAX_TRACES Then
        mTruncated = True
    Else
        out.Add tr
    End If
End Sub

Private Sub ApplyAction_(ByVal node As OrderedDict, ByVal stack As Collection, ByVal tr As OrderedDict, _
                         ByVal out As Collection, ByVal secLabels As OrderedDict)
    Static rxPerf As Object, rxThru As Object, rxCall As Object, rxMove As Object
    Static rxComp As Object, rxStr As Object, rxInit As Object, rxAcc As Object
    If rxPerf Is Nothing Then
        Set rxPerf = CreateObject("VBScript.RegExp"): rxPerf.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9-]*)$": rxPerf.IgnoreCase = False
        Set rxThru = CreateObject("VBScript.RegExp"): rxThru.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9-]*)\s+THR(U|OUGH)\s+([A-Z0-9][A-Z0-9-]*)": rxThru.IgnoreCase = False
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

    ' PERFORM x THRU y
    Set m = rxThru.Execute(lbl)
    If m.Count > 0 Then
        PerformInto_ m.Item(0).SubMatches(0), m.Item(0).SubMatches(2), node, stack, tr, out, secLabels
        Exit Sub
    End If

    ' PERFORM x
    Set m = rxPerf.Execute(lbl)
    If m.Count > 0 Then
        PerformInto_ m.Item(0).SubMatches(0), "", node, stack, tr, out, secLabels
        Exit Sub
    End If

    ' CALL
    Set m = rxCall.Execute(lbl)
    If m.Count > 0 Then
        Dim tgt As String, params As String
        tgt = m.Item(0).SubMatches(0)
        params = m.Item(0).SubMatches(2)
        tr.Add "Calls", Cons_(tr.Item("Calls"), tgt)
        AddEvent_ tr, "call", lbl, ln
        ' remember the first call site for synthesized failure cases
        If Not mSynth.Exists(tgt) Then
            Dim sn As OrderedDict
            Set sn = New OrderedDict
            sn.Add "Trace", CloneTrace_(tr)
            sn.Add "Line", ln
            sn.Add "Target", tgt
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
                If GetLiteral_(srcTxt, lit) Then tr.Item("Env").Add dst, lit
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

    ' EXEC and anything else: generic event
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
        AddCapped_ out, sb
    Next sb
End Sub

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

    Dim cs As Collection, w As OrderedDict, lit As String, isOther As Boolean
    Set cs = node.Item("cases")
    Dim anyTaken As Boolean, litAll As Boolean
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

    Dim c1 As OrderedDict, seed As Collection, subs As Collection, sb As OrderedDict
    For wi = 1 To cs.Count
        Set w = cs(wi)
        isOther = (CStr(w.Item("condition")) = "OTHER")
        Dim takeIt As Boolean
        takeIt = True
        If litAll Then
            If matchedIdx > 0 Then
                takeIt = (wi = matchedIdx)
            Else
                takeIt = isOther   ' no literal matched: only OTHER (or skip)
            End If
        End If
        If takeIt Then
            Set c1 = CloneTrace_(tr)
            AddArmEvent_ c1, CStr(w.Item("id")), "WHEN", expr & " = " & CStr(w.Item("condition")), CLng(w.Item("startLine"))
            Set seed = New Collection
            seed.Add c1
            Set subs = ApplyNodes_(w.Item("children"), stack, seed, secLabels)
            For Each sb In subs: AddCapped_ out, sb: Next sb
        End If
    Next wi

    ' implicit skip arm when no WHEN OTHER exists
    If Not hasOther Then
        Dim skipIt As Boolean
        skipIt = True
        If litAll And matchedIdx > 0 Then skipIt = False
        If skipIt Then
            Set c1 = CloneTrace_(tr)
            AddArmEvent_ c1, CStr(node.Item("id")) & ":skip", "SKIP", expr & " = (no match)", CLng(node.Item("startLine"))
            AddCapped_ out, c1
        End If
    End If
End Sub

Private Sub ApplySearch_(ByVal node As OrderedDict, ByVal stack As Collection, ByVal tr As OrderedDict, _
                         ByVal out As Collection, ByVal secLabels As OrderedDict)
    Dim c1 As OrderedDict, seed As Collection, subs As Collection, sb As OrderedDict
    ' AT END arm (or implicit skip when absent)
    Set c1 = CloneTrace_(tr)
    If Not IsNull(node.Item("atEndLine")) Then
        AddArmEvent_ c1, CStr(node.Item("id")) & ":atend", "AT END", "SEARCH " & CStr(node.Item("tableExpr")), CLng(node.Item("startLine"))
        Set seed = New Collection
        seed.Add c1
        Set subs = ApplyNodes_(node.Item("atEndChildren"), stack, seed, secLabels)
        For Each sb In subs: AddCapped_ out, sb: Next sb
    Else
        AddArmEvent_ c1, CStr(node.Item("id")) & ":skip", "AT END(skip)", "SEARCH " & CStr(node.Item("tableExpr")), CLng(node.Item("startLine"))
        AddCapped_ out, c1
    End If
    ' WHEN arms
    Dim cs As Collection, wi As Long, w As OrderedDict
    Set cs = node.Item("cases")
    For wi = 1 To cs.Count
        Set w = cs(wi)
        Set c1 = CloneTrace_(tr)
        AddArmEvent_ c1, CStr(w.Item("id")), "WHEN", CStr(w.Item("condition")), CLng(w.Item("startLine"))
        Set seed = New Collection
        seed.Add c1
        Set subs = ApplyNodes_(w.Item("children"), stack, seed, secLabels)
        For Each sb In subs: AddCapped_ out, sb: Next sb
    Next wi
End Sub

'======================================================================
' Arms catalog
'======================================================================
Private Sub CollectArms_(ByVal nodes As Collection, ByVal arms As Collection)
    Dim n As OrderedDict, t As String
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "if" Then
            AddArm_ arms, CStr(n.Item("id")) & ":then", CLng(n.Item("startLine")), "IF " & CStr(n.Item("condition")) & " [THEN]"
            AddArm_ arms, CStr(n.Item("id")) & ":else", CLng(n.Item("startLine")), "IF " & CStr(n.Item("condition")) & " [ELSE]"
            CollectArms_ n.Item("thenChildren"), arms
            CollectArms_ n.Item("elseChildren"), arms
        ElseIf t = "evaluate" Then
            Dim cs As Collection, wi As Long, w As OrderedDict, hasOther As Boolean
            Set cs = n.Item("cases")
            hasOther = False
            For wi = 1 To cs.Count
                Set w = cs(wi)
                AddArm_ arms, CStr(w.Item("id")), CLng(w.Item("startLine")), "EVALUATE " & CStr(n.Item("expression")) & " WHEN " & CStr(w.Item("condition"))
                If CStr(w.Item("condition")) = "OTHER" Then hasOther = True
                CollectArms_ w.Item("children"), arms
            Next wi
            If Not hasOther Then
                AddArm_ arms, CStr(n.Item("id")) & ":skip", CLng(n.Item("startLine")), "EVALUATE " & CStr(n.Item("expression")) & " [no-match skip]"
            End If
        ElseIf t = "search" Then
            If Not IsNull(n.Item("atEndLine")) Then
                AddArm_ arms, CStr(n.Item("id")) & ":atend", CLng(n.Item("startLine")), "SEARCH " & CStr(n.Item("tableExpr")) & " [AT END]"
            Else
                AddArm_ arms, CStr(n.Item("id")) & ":skip", CLng(n.Item("startLine")), "SEARCH " & CStr(n.Item("tableExpr")) & " [AT END skip]"
            End If
            Dim sc As Collection, si As Long, sw As OrderedDict
            Set sc = n.Item("cases")
            For si = 1 To sc.Count
                Set sw = sc(si)
                AddArm_ arms, CStr(sw.Item("id")), CLng(sw.Item("startLine")), "SEARCH WHEN " & CStr(sw.Item("condition"))
                CollectArms_ sw.Item("children"), arms
            Next si
        End If
    Next n
End Sub

Private Sub AddArm_(ByVal arms As Collection, ByVal token As String, ByVal lineNo As Long, ByVal disp As String)
    Dim a As OrderedDict
    Set a = New OrderedDict
    a.Add "Token", token
    a.Add "Line", lineNo
    a.Add "Disp", disp
    arms.Add a
End Sub

'======================================================================
' Case selection (C1 greedy + abend grouping + synthesized failures)
'======================================================================
Private Function SelectCases_(ByVal normals As Collection, ByVal abends As Collection, _
                              ByVal arms As Collection) As Collection
    Dim cases As Collection
    Set cases = New Collection

    ' materialize each normal trace's arms once (the greedy loop below
    ' iterates them repeatedly; cons heads are walk-once structures)
    Dim coverable As OrderedDict, tr As OrderedDict, v As Variant
    For Each tr In normals
        tr.Add "ArmsL", ConsToList_(tr.Item("Arms"))
    Next tr

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

    ' synthesized call-failure stubs
    Dim synths As Collection
    Set synths = New Collection
    Dim sk As Collection, sv As Variant, sn As OrderedDict, st As OrderedDict
    Set sk = mSynth.Keys
    For Each sv In sk
        Set sn = mSynth.Item(CStr(sv))
        Set st = sn.Item("Trace")
        st.Add "Term", "synth:" & CStr(sn.Item("Target"))
        st.Add "TriggerLine", CLng(sn.Item("Line"))
        AddEvent_ st, "term", "CALL " & CStr(sn.Item("Target")) & " -> ABNORMAL (synthesized)", CLng(sn.Item("Line"))
        synths.Add st
    Next sv

    ' abnormal = code abends + synths, ordered by trigger line
    Dim abnormal As Collection
    Set abnormal = New Collection
    Dim gk As Collection, gv As Variant
    Set gk = groups.Keys
    For Each gv In gk
        Set tr = groups.Item(CStr(gv))
        If CLng(tr.Item("TriggerLine")) = 0 Then
            ' trigger = line of the terminating event (= the cons head)
            Dim eh As ConsList
            Set eh = tr.Item("Events")
            If Not eh Is Nothing Then tr.Add "TriggerLine", CLng(eh.V.Item("Line"))
        End If
        abnormal.Add tr
    Next gv
    For Each tr In synths
        abnormal.Add tr
    Next tr
    Set abnormal = SortByTrigger_(abnormal)

    ' assemble case records
    Dim serial As Long, normSerial As Long, abSerial As Long
    serial = 0
    For Each tr In picked
        serial = serial + 1
        normSerial = normSerial + 1
        cases.Add BuildCase_(tr, "TC" & serial, "normal", normSerial)
    Next tr
    For Each tr In abnormal
        serial = serial + 1
        abSerial = abSerial + 1
        If Left$(CStr(tr.Item("Term")), 6) = "synth:" Then
            cases.Add BuildCase_(tr, "TC" & serial, "synth", abSerial)
        Else
            cases.Add BuildCase_(tr, "TC" & serial, "abend", abSerial)
        End If
    Next tr

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
    If tr.Exists("ArmsL") Then
        c.Add "arms", tr.Item("ArmsL")
    Else
        c.Add "arms", ConsToList_(tr.Item("Arms"))
    End If
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
