Attribute VB_Name = "CobolExecTree"
' CobolExecTree - render the ロジック階層 tree in COBOL EXECUTION ORDER.
'
' Instead of source order, the tree starts at the PROCEDURE DIVISION entry
' section and follows PERFORM (and PERFORM .. THRU ..) calls, inlining each
' performed paragraph/section at its call site (recursively), with cycle
' detection. Subroutine paragraphs therefore appear only where they are
' performed - the sequential body of a routine is capped at its first
' PERFORM-target sub-paragraph, so nothing is rendered twice.
'
' The column layout is identical to CobolLogicViewer.RenderHierarchy
' (A = tree, B = 開始行 hyperlink, C = 種別), so CobolTcMark's test-case
' marking and column-A highlight keep working unchanged. Approach mirrors the
' sibling tool F:\Dev\cobol\cobol-logic-hierarchy (also Pure VBA).

Option Explicit

Private Const C_HEADER  As Long = 14474460
Private Const C_BRANCH  As Long = 16314338
Private Const C_SECTION As Long = 13168895
' Hidden column carrying each row's branch context (the pipe-joined branchId
' tokens of the arms it is nested under). CobolTcMark reads it to highlight the
' lines a selected test case actually executes.
Private Const COL_CTX As Long = 7
' Sentinel context for rows in entry-unreached sections: a token no enumerated
' path's branchIds contains, so coverage never paints them as executed.
Private Const UNREACHED_CTX As String = "__unreached__"
Private Const MAX_TREE_ROWS As Long = 8000   ' render guard for heavily re-PERFORMed programs
Private mOverflow As Boolean

Private mWs As Worksheet
Private mRow As Long
Private mNodes As Object        ' root("rootNodes")  (1-based Collection)
Private mOwners As Collection   ' OrderedDict {name, line, kind, ownerEnd, secEnd}, sorted by line
Private mCut As OrderedDict      ' names targeted by a plain PERFORM -> True
Private mVisited As OrderedDict  ' section/para names already rendered

Public Sub RenderExecHierarchy(ByVal ws As Worksheet, ByVal root As Object)
    ws.Cells.Clear
    Set mWs = ws

    Dim s As Object
    Set s = root("summary")
    ws.Range("A1").Value = "COBOLロジック階層（実行順展開）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "プログラム名": ws.Range("B3").Value = s("programName")
    ws.Range("A4").Value = "行数": ws.Range("B4").Value = s("lines")
    ws.Range("A5").Value = "分岐": ws.Range("B5").Value = s("branchCount")
    ws.Range("A6").Value = "パス数": ws.Range("B6").Value = s("pathCount")
    ws.Range("A7").Value = "プレフィックス形式": ws.Range("B7").Value = s("prefixStyle") & " (検出: " & s("prefixDetected") & ")"
    ws.Range("A8").Value = "警告件数": ws.Range("B8").Value = root("warnings").Count

    Dim hdr As Long
    hdr = 10
    ws.Cells(hdr, 1).Value = "階層ツリー"
    ws.Cells(hdr, 2).Value = "開始行"
    ws.Cells(hdr, 3).Value = "種別"
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 3)).Font.Bold = True
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 3)).Interior.Color = C_HEADER

    mRow = hdr + 1
    Dim treeStart As Long
    treeStart = mRow

    Set mNodes = root("rootNodes")
    BuildOwners_ root
    BuildCut_ root
    Set mVisited = New OrderedDict
    mOverflow = False

    If mNodes.Count > 0 And mOwners.Count > 0 Then
        Dim firstStmt As Long, maxStmt As Long
        ComputeStmtRange_ firstStmt, maxStmt
        Dim entry As OrderedDict
        Set entry = FindProcEntrySection_(firstStmt)
        Dim pathStack As Collection
        Set pathStack = New Collection
        If Not entry Is Nothing Then
            ' normal case: a procedure SECTION owns the entry point
            EmitSectionHeader_ CStr(entry.Item("name"))
            mVisited.Add CStr(entry.Item("name")), True
            pathStack.Add CStr(entry.Item("name"))
            RenderNodeList_ CLng(entry.Item("line")), CLng(entry.Item("secEnd")), "", pathStack, ""
            pathStack.Remove pathStack.Count
            RenderUnreached_ entry, pathStack
        Else
            ' fallback: paragraph-only program (no procedure SECTION). Render the
            ' main flow directly from the first statement; PERFORMs still inline.
            RenderNodeList_ firstStmt, maxStmt, "", pathStack, ""
        End If
    End If

    Dim treeEnd As Long
    treeEnd = mRow - 1
    If mOverflow Then
        mWs.Cells(mRow, 1).Value = "※ 展開行数が上限(" & MAX_TREE_ROWS & "行)を超えたため以降を省略しました（PERFORM 多用プログラム）"
        mWs.Cells(mRow, 1).Font.Color = RGB(192, 0, 0)
        mWs.Cells(mRow, COL_CTX).Value = UNREACHED_CTX   ' keep coverage marking off this row
        mRow = mRow + 1
    End If
    If treeEnd >= treeStart Then
        ' HYPERLINK formulas do not get the built-in link style - restore it once
        With mWs.Range(mWs.Cells(treeStart, 2), mWs.Cells(treeEnd, 2)).Font
            .Color = RGB(5, 99, 193)
            .Underline = True
        End With
    End If
    If treeEnd >= treeStart Then
        ws.Range(ws.Cells(treeStart, 1), ws.Cells(treeEnd, 1)).Font.Name = "MS Gothic"
    End If
    ws.Columns("A:C").AutoFit
    If ws.Columns("A").ColumnWidth > 90 Then ws.Columns("A").ColumnWidth = 90
    ws.Columns(COL_CTX).Hidden = True
End Sub

' ---- structure helpers ------------------------------------------------ '

Private Sub BuildOwners_(ByVal root As Object)
    Dim raw As Collection
    Set raw = New Collection
    Dim ps As Object
    Set ps = root("programStructure")
    Dim e As Object
    For Each e In ps("sections")
        AddOwner_ raw, CStr(e("name")), CLng(e("line")), "section"
    Next e
    For Each e In ps("paragraphs")
        AddOwner_ raw, CStr(e("name")), CLng(e("line")), "para"
    Next e

    Set mOwners = SortOwnersByLine_(raw)

    Dim i As Long, o As OrderedDict, j As Long, oe As Long, se As Long
    For i = 1 To mOwners.Count
        Set o = mOwners(i)
        oe = 999999
        If i < mOwners.Count Then oe = CLng(mOwners(i + 1).Item("line")) - 1
        se = 999999
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

Private Sub AddOwner_(ByVal c As Collection, ByVal nm As String, ByVal ln As Long, ByVal kind As String)
    Dim o As OrderedDict
    Set o = New OrderedDict
    o.Add "name", UCase$(nm)
    o.Add "line", ln
    o.Add "kind", kind
    c.Add o
End Sub

Private Function SortOwnersByLine_(ByVal raw As Collection) As Collection
    ' selection sort into a new collection (owner count is small)
    Dim result As Collection
    Set result = New Collection
    If raw.Count = 0 Then
        Set SortOwnersByLine_ = result
        Exit Function
    End If
    Dim used() As Boolean
    ReDim used(1 To raw.Count)
    Dim k As Long, i As Long, best As Long, bestLine As Long
    For k = 1 To raw.Count
        best = 0
        bestLine = 2000000000
        For i = 1 To raw.Count
            If Not used(i) Then
                If CLng(raw(i).Item("line")) < bestLine Then
                    bestLine = CLng(raw(i).Item("line"))
                    best = i
                End If
            End If
        Next i
        used(best) = True
        result.Add raw(best)
    Next k
    Set SortOwnersByLine_ = result
End Function

Private Sub BuildCut_(ByVal root As Object)
    Set mCut = New OrderedDict
    Dim cg As Object
    On Error Resume Next
    Set cg = root("callGraph")
    On Error GoTo 0
    If cg Is Nothing Then Exit Sub
    Dim edge As Object
    For Each edge In cg("edges")
        If CStr(edge("kind")) = "perform" Then
            If Not mCut.Exists(CStr(edge("to"))) Then mCut.Add CStr(edge("to")), True
        End If
    Next edge
End Sub

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

Private Sub ComputeStmtRange_(ByRef firstStmt As Long, ByRef maxStmt As Long)
    firstStmt = 2000000000
    maxStmt = 0
    Dim n As Object, i As Long, ln As Long
    For i = 1 To mNodes.Count
        Set n = mNodes(i)
        ln = CLng(n("startLine"))
        If ln < firstStmt Then firstStmt = ln
        If ln > maxStmt Then maxStmt = ln
    Next i
End Sub

' The PROCEDURE-DIVISION entry section = the last SECTION defined at or before
' the first statement whose name is NOT a DATA/ENVIRONMENT division section
' (WORKING-STORAGE etc.). Returns Nothing for paragraph-only programs, in which
' case the caller renders the main flow directly.
Private Function FindProcEntrySection_(ByVal firstStmt As Long) As OrderedDict
    Dim o As OrderedDict, found As OrderedDict
    Set found = Nothing
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" And CLng(o.Item("line")) <= firstStmt Then
            If Not IsDataSection_(CStr(o.Item("name"))) Then Set found = o
        End If
    Next o
    Set FindProcEntrySection_ = found
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

' Highest line to render sequentially for a routine starting at lo: stop just
' before the first PERFORM-target sub-paragraph (so subroutines are only shown
' inlined at their call site, never duplicated as fall-through).
Private Function CapHi_(ByVal lo As Long, ByVal hi As Long) As Long
    Dim o As OrderedDict, c As Long
    c = 2000000000
    For Each o In mOwners
        If mCut.Exists(CStr(o.Item("name"))) Then
            Dim ln As Long
            ln = CLng(o.Item("line"))
            If ln > lo And ln <= hi And ln < c Then c = ln
        End If
    Next o
    If c < 2000000000 Then
        CapHi_ = c - 1
    Else
        CapHi_ = hi
    End If
End Function

Private Function OnPath_(ByVal stack As Collection, ByVal nm As String) As Boolean
    Dim v As Variant
    For Each v In stack
        If CStr(v) = nm Then
            OnPath_ = True
            Exit Function
        End If
    Next v
    OnPath_ = False
End Function

Private Function ResolvePerform_(ByVal label As String) As OrderedDict
    Set ResolvePerform_ = Nothing
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9_-]+)(\s+THR(U|OUGH)\s+([A-Z0-9][A-Z0-9_-]+))?"
        rx.IgnoreCase = False
    End If
    Dim m As Object
    Set m = rx.Execute(UCase$(label))
    If m.Count = 0 Then Exit Function
    Dim x As String, y As String
    x = m.Item(0).SubMatches(0)
    y = m.Item(0).SubMatches(3)
    Dim ox As OrderedDict
    Set ox = OwnerByName_(x)
    If ox Is Nothing Then Exit Function
    Dim hi As Long
    If CStr(ox.Item("kind")) = "section" Then hi = CLng(ox.Item("secEnd")) Else hi = CLng(ox.Item("ownerEnd"))
    If Len(y) > 0 Then
        Dim oy As OrderedDict
        Set oy = OwnerByName_(y)
        If Not oy Is Nothing Then
            If CStr(oy.Item("kind")) = "section" Then hi = CLng(oy.Item("secEnd")) Else hi = CLng(oy.Item("ownerEnd"))
        End If
    End If
    Dim r As OrderedDict
    Set r = New OrderedDict
    r.Add "name", x
    r.Add "lo", CLng(ox.Item("line"))
    r.Add "hi", hi
    r.Add "kind", CStr(ox.Item("kind"))
    Set ResolvePerform_ = r
End Function

Private Function NodesInRange_(ByVal lo As Long, ByVal hi As Long) As Collection
    Dim c As Collection
    Set c = New Collection
    Dim i As Long, n As Object, ln As Long
    For i = 1 To mNodes.Count
        Set n = mNodes(i)
        ln = CLng(n("startLine"))
        If ln >= lo And ln <= hi Then c.Add n
    Next i
    Set NodesInRange_ = c
End Function

' ---- rendering -------------------------------------------------------- '

Private Sub RenderNodeList_(ByVal lo As Long, ByVal hi As Long, ByVal prefix As String, _
                            ByVal pathStack As Collection, ByVal context As String)
    Dim eh As Long
    eh = CapHi_(lo, hi)
    Dim list As Collection
    Set list = NodesInRange_(lo, eh)
    Dim i As Long
    For i = 1 To list.Count
        RenderNode_ list(i), prefix, (i = list.Count), pathStack, context
    Next i
End Sub

Private Sub RenderNode_(ByVal node As Object, ByVal prefix As String, ByVal isLast As Boolean, _
                        ByVal pathStack As Collection, ByVal context As String)
    Dim connector As String, childPrefix As String
    If isLast Then
        connector = prefix & ChrW$(&H2514) & ChrW$(&H2500) & " "
        childPrefix = prefix & "    "
    Else
        connector = prefix & ChrW$(&H251C) & ChrW$(&H2500) & " "
        childPrefix = prefix & ChrW$(&H2502) & "   "
    End If

    Dim t As String, label As String, kind As String, nid As String
    t = node("type")
    nid = CStr(node("id"))
    Select Case t
        Case "if"
            label = "IF " & node("condition"): kind = "IF"
        Case "evaluate"
            label = "EVALUATE " & node("expression"): kind = "EVALUATE"
        Case "search"
            label = "SEARCH " & IIf(node("isAll"), "ALL ", "") & node("tableExpr"): kind = "SEARCH"
        Case "when"
            label = "WHEN " & node("condition"): kind = "WHEN"
        Case "action"
            label = node("label"): kind = "ACTION"
        Case Else
            label = t: kind = t
    End Select

    ' A WHEN row is executed only when its arm is taken, so the row itself must
    ' carry its own when-id token (not just its children). Other node types keep
    ' the parent context (they execute whenever reached).
    Dim rowCtx As String
    rowCtx = context
    If t = "when" Then rowCtx = AppendCtx_(context, nid)
    EmitRow_ connector & label, CLng(node("startLine")), kind, (t <> "action"), rowCtx

    If t = "action" Then
        Dim tgt As OrderedDict
        Set tgt = ResolvePerform_(CStr(node("label")))
        If Not tgt Is Nothing Then
            If OnPath_(pathStack, CStr(tgt.Item("name"))) Then
                EmitRow_ childPrefix & ChrW$(&H2514) & ChrW$(&H2500) & " （再帰: " & tgt.Item("name") & " 展開済）", 0, "", False, context
            Else
                If Not mVisited.Exists(CStr(tgt.Item("name"))) Then mVisited.Add CStr(tgt.Item("name")), True
                pathStack.Add CStr(tgt.Item("name"))
                ' banner marking where the PERFORM target is inline-expanded
                EmitBanner_ childPrefix, tgt, context
                ' inlined body shares this PERFORM's branch context
                RenderNodeList_ CLng(tgt.Item("lo")), CLng(tgt.Item("hi")), childPrefix, pathStack, context
                pathStack.Remove pathStack.Count
            End If
        End If
        Exit Sub
    End If

    Select Case t
        Case "if"
            Dim elseCount As Long
            elseCount = node("elseChildren").Count
            RenderBranchGroup_ "[THEN]", node("thenChildren"), childPrefix, (elseCount = 0), pathStack, AppendCtx_(context, nid & ":then")
            If elseCount > 0 Then
                RenderBranchGroup_ "[ELSE]", node("elseChildren"), childPrefix, True, pathStack, AppendCtx_(context, nid & ":else")
            End If
        Case "evaluate"
            Dim cs As Object, ci As Long
            Set cs = node("cases")
            For ci = 1 To cs.Count
                RenderNode_ cs(ci), childPrefix, (ci = cs.Count), pathStack, context
            Next ci
        Case "search"
            Dim atEnd As Object, sc As Object, si As Long
            Set atEnd = node("atEndChildren")
            Set sc = node("cases")
            If atEnd.Count > 0 Then
                RenderBranchGroup_ "[AT END]", atEnd, childPrefix, (sc.Count = 0), pathStack, AppendCtx_(context, nid & ":atend")
            End If
            For si = 1 To sc.Count
                RenderNode_ sc(si), childPrefix, (si = sc.Count), pathStack, context
            Next si
        Case "when"
            ' children of a WHEN share the WHEN row's own context (parent + when-id)
            Dim wc As Object, wi As Long
            Set wc = node("children")
            For wi = 1 To wc.Count
                RenderNode_ wc(wi), childPrefix, (wi = wc.Count), pathStack, rowCtx
            Next wi
    End Select
End Sub

Private Sub RenderBranchGroup_(ByVal groupLabel As String, ByVal children As Object, _
                               ByVal prefix As String, ByVal isLast As Boolean, _
                               ByVal pathStack As Collection, ByVal context As String)
    Dim connector As String, childPrefix As String
    If isLast Then
        connector = prefix & ChrW$(&H2514) & ChrW$(&H2500) & " "
        childPrefix = prefix & "    "
    Else
        connector = prefix & ChrW$(&H251C) & ChrW$(&H2500) & " "
        childPrefix = prefix & ChrW$(&H2502) & "   "
    End If
    EmitRow_ connector & groupLabel, 0, "BRANCH", True, context
    Dim i As Long
    For i = 1 To children.Count
        RenderNode_ children(i), childPrefix, (i = children.Count), pathStack, context
    Next i
End Sub

Private Sub RenderUnreached_(ByVal entry As OrderedDict, ByVal pathStack As Collection)
    Dim entryLine As Long
    entryLine = 0
    If Not entry Is Nothing Then entryLine = CLng(entry.Item("line"))
    Dim o As OrderedDict, firstUn As Boolean
    firstUn = True
    For Each o In mOwners
        If CStr(o.Item("kind")) = "section" And CLng(o.Item("line")) >= entryLine Then
            If Not mVisited.Exists(CStr(o.Item("name"))) Then
                ' only render sections that actually contain statements
                Dim probe As Collection
                Set probe = NodesInRange_(CLng(o.Item("line")), CLng(o.Item("secEnd")))
                If probe.Count > 0 Then
                    If firstUn Then
                        EmitRow_ "── （entry から未到達のセクション） ──", 0, "", False, UNREACHED_CTX
                        firstUn = False
                    End If
                    EmitSectionHeader_ CStr(o.Item("name"))
                    mVisited.Add CStr(o.Item("name")), True
                    pathStack.Add CStr(o.Item("name"))
                    RenderNodeList_ CLng(o.Item("line")), CLng(o.Item("secEnd")), "", pathStack, UNREACHED_CTX
                    pathStack.Remove pathStack.Count
                End If
            End If
        End If
    Next o
End Sub

' Banner row at each PERFORM inline-expansion site, indented to align with the
' expanded body, so expansion boundaries are easy to spot (team request).
' Reuses EmitRow_ for the line-number hyperlink, then restyles the text.
Private Sub EmitBanner_(ByVal prefix As String, ByVal tgt As OrderedDict, ByVal context As String)
    Dim sfx As String
    If CStr(tgt.Item("kind")) = "section" Then sfx = " SECTION" Else sfx = ""
    Dim txt As String
    txt = prefix & ChrW$(&H25A0) & "--------------- " & CStr(tgt.Item("name")) & sfx & " ---------------"
    EmitRow_ txt, CLng(tgt.Item("lo")), "", False, context
    If mOverflow Then Exit Sub   ' row was suppressed - do not restyle the previous row
    mWs.Cells(mRow - 1, 1).Font.Bold = True
End Sub
Private Sub EmitSectionHeader_(ByVal nm As String)
    If mRow > MAX_TREE_ROWS Then
        mOverflow = True
        Exit Sub
    End If
    mWs.Cells(mRow, 1).Value = ChrW$(&H25A0) & " " & nm & " SECTION"
    mWs.Cells(mRow, 1).Font.Bold = True
    mWs.Range(mWs.Cells(mRow, 1), mWs.Cells(mRow, 3)).Interior.Color = C_SECTION
    mRow = mRow + 1
End Sub

Private Function AppendCtx_(ByVal ctx As String, ByVal token As String) As String
    If Len(ctx) = 0 Then
        AppendCtx_ = token
    Else
        AppendCtx_ = ctx & "|" & token
    End If
End Function

Private Sub EmitRow_(ByVal text As String, ByVal lineNo As Long, ByVal kind As String, _
                     ByVal branchColor As Boolean, ByVal context As String)
    If mRow > MAX_TREE_ROWS Then
        mOverflow = True
        Exit Sub
    End If
    mWs.Cells(mRow, 1).Value = text
    If lineNo > 0 Then
        On Error Resume Next
        ' HYPERLINK formula instead of Hyperlinks.Add: same click-to-jump,
        ' but a plain value write (~100x faster on 1000+ row trees).
        mWs.Cells(mRow, 2).Formula = "=HYPERLINK(""#'COBOLソース'!A" & (lineNo + 3) & """," & lineNo & ")"
        If Err.Number <> 0 Then
            mWs.Cells(mRow, 2).Value = lineNo
            Err.Clear
        End If
        On Error GoTo 0
    End If
    If Len(kind) > 0 Then mWs.Cells(mRow, 3).Value = kind
    If branchColor Then
        mWs.Range(mWs.Cells(mRow, 1), mWs.Cells(mRow, 3)).Interior.Color = C_BRANCH
    End If
    ' record the branch context (used by CobolTcMark for coverage highlighting);
    ' tokens always start with a letter (if-/when-/search-), so no '=' formula risk
    If Len(context) > 0 Then mWs.Cells(mRow, COL_CTX).Value = context
    mRow = mRow + 1
End Sub
