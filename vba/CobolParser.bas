Attribute VB_Name = "CobolParser"
' CobolParser - COBOL logic analyzer engine.
' Builds an AST from COBOL source, enumerates execution paths, and computes
' branch coverage and call relationships. Result is an OrderedDict tree
' matching the JSON schema consumed by CobolLogicViewer.

Option Explicit

Public Const PARSER_VERSION As String = "0.4.0"

' Cap on enumerated execution paths. Above this we stop expanding and set
' pathTruncated=true so the report still emits something useful.
Public Const MAX_PATH_STATES As Long = 200

' Module-level state populated during analysis. Reset by Analyze_Full / _Phase2.
Private mUnclosedFrames As Long
Private mPathTruncated As Boolean

Public Property Get UnclosedFrames() As Long
    UnclosedFrames = mUnclosedFrames
End Property

Public Property Get PathTruncated() As Boolean
    PathTruncated = mPathTruncated
End Property

'==============================================================================
' String helpers
'==============================================================================

Public Function Convert_CollapseSpaces(ByVal txt As String) As String
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "\s+"
        rx.Global = True
    End If
    Convert_CollapseSpaces = Trim$(rx.Replace(txt, " "))
End Function

Public Function Convert_StripTrailingPeriod(ByVal txt As String) As String
    Dim t As String
    t = Trim$(txt)
    If Len(t) > 0 Then
        If Right$(t, 1) = "." Then t = Left$(t, Len(t) - 1)
    End If
    Convert_StripTrailingPeriod = Trim$(t)
End Function

Public Function Convert_NormalizeCondition(ByVal cond As String) As String
    Convert_NormalizeCondition = Convert_CollapseSpaces(Convert_StripTrailingPeriod(cond))
End Function

Public Function Convert_InvertCondition(ByVal cond As String) As String
    If UCase$(Trim$(cond)) = "OTHER" Then
        Convert_InvertCondition = "NOT (OTHER)"
    Else
        Convert_InvertCondition = "NOT (" & cond & ")"
    End If
End Function

'==============================================================================
' Line normalization + prefix auto-detection
'==============================================================================

' Returns OrderedDict { Lines, PrefixDetected, PrefixStyle, PrefixRatio }
' where Lines is a Collection of OrderedDict { Number, Raw, Text }.
Public Function Get_NormalizedCobolLines(ByVal source As String, Optional ByVal forcePrefix As String = "") As OrderedDict
    Dim rawLines() As String
    rawLines = Split(Replace(source, Chr$(13), ""), Chr$(10))

    Dim rxPrefixNum As Object, rxPrefixSep As Object, rxStandardHead As Object, rxComment As Object
    Set rxPrefixNum = CreateObject("VBScript.RegExp")
    rxPrefixNum.Pattern = "^\d{6}$"
    Set rxPrefixSep = CreateObject("VBScript.RegExp")
    rxPrefixSep.Pattern = "[\s*/]"
    Set rxStandardHead = CreateObject("VBScript.RegExp")
    rxStandardHead.Pattern = "^\s*\d{0,6}\s"
    Set rxComment = CreateObject("VBScript.RegExp")
    rxComment.Pattern = "^[*/]"

    Dim nonBlank As Long, prefixHit As Long
    Dim i As Long, rl As String
    For i = LBound(rawLines) To UBound(rawLines)
        rl = rawLines(i)
        If Len(Trim$(rl)) > 0 Then
            nonBlank = nonBlank + 1
            If Len(rl) >= 7 Then
                If rxPrefixNum.Test(Left$(rl, 6)) And rxPrefixSep.Test(Mid$(rl, 7, 1)) Then
                    prefixHit = prefixHit + 1
                End If
            End If
        End If
    Next i

    Dim ratio As Double
    If nonBlank > 0 Then ratio = prefixHit / nonBlank

    Dim style As String, detected As Boolean
    Select Case LCase$(forcePrefix)
        Case "prefixed":  style = "prefixed":  detected = True
        Case "standard":  style = "standard":  detected = False
        Case "none":      style = "none":      detected = False
        Case Else
            If ratio >= 0.6 Then
                style = "prefixed":  detected = True
            Else
                style = "standard":  detected = False
            End If
    End Select

    Dim lines As Collection
    Set lines = New Collection

    Dim idx As Long, stripped As String, headStandard As String, txt As String
    Dim entry As OrderedDict
    For i = LBound(rawLines) To UBound(rawLines)
        rl = rawLines(i)
        idx = idx + 1
        If style = "prefixed" Then
            If Len(rl) >= 6 Then stripped = Mid$(rl, 7) Else stripped = rl
        ElseIf style = "none" Then
            stripped = rl
        Else
            If Len(rl) >= 8 Then headStandard = Left$(rl, 8) Else headStandard = rl
            If Len(rl) > 6 And rxStandardHead.Test(headStandard) Then
                stripped = Mid$(rl, 7)
            Else
                stripped = rl
            End If
        End If
        txt = Convert_CollapseSpaces(stripped)
        If txt <> "" And Not rxComment.Test(txt) Then
            Set entry = New OrderedDict
            entry.Add "Number", idx
            entry.Add "Raw", rl
            entry.Add "Text", txt
            lines.Add entry
        End If
    Next i

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "Lines", lines
    result.Add "PrefixDetected", detected
    result.Add "PrefixStyle", style
    result.Add "PrefixRatio", Round(ratio, 3)
    Set Get_NormalizedCobolLines = result
End Function

'==============================================================================
' Identifier / structure extraction
'==============================================================================

Public Function Get_ProgramName(ByVal lines As Collection) As String
    Dim rxPid As Object, rxSec As Object, entry As OrderedDict, m As Object
    Set rxPid = CreateObject("VBScript.RegExp")
    rxPid.Pattern = "^PROGRAM-ID\.\s*([A-Z0-9][A-Z0-9_-]*)"
    rxPid.IgnoreCase = False
    Set rxSec = CreateObject("VBScript.RegExp")
    rxSec.Pattern = "^([A-Z0-9][A-Z0-9_-]+)\s+SECTION$"
    rxSec.IgnoreCase = False

    For Each entry In lines
        Set m = rxPid.Execute(entry.Item("Text"))
        If m.Count > 0 Then
            Get_ProgramName = UCase$(m.Item(0).SubMatches(0))
            Exit Function
        End If
    Next entry
    For Each entry In lines
        Set m = rxSec.Execute(entry.Item("Text"))
        If m.Count > 0 Then
            Get_ProgramName = UCase$(m.Item(0).SubMatches(0))
            Exit Function
        End If
    Next entry
    Get_ProgramName = "(NO-NAME)"
End Function

Public Function Test_ParagraphName(ByVal name As String) As Boolean
    Dim excluded As Variant
    excluded = Array("IDENTIFICATION", "ENVIRONMENT", "DATA", "PROCEDURE", "PROGRAM-ID", _
                     "THEN", "ELSE", "END-IF", "END-EVALUATE", "END-SEARCH", "END-PERFORM", _
                     "END-READ", "END-CALL", "EXIT", "CONTINUE", "GOBACK")
    Dim upName As String, i As Long
    upName = UCase$(name)
    For i = LBound(excluded) To UBound(excluded)
        If excluded(i) = upName Then
            Test_ParagraphName = False
            Exit Function
        End If
    Next i
    Test_ParagraphName = True
End Function

' Returns OrderedDict { sections, paragraphs, performThruRanges }.
Public Function Get_ProgramStructure(ByVal lines As Collection) As OrderedDict
    Dim sections As Collection, paragraphs As Collection, performThru As Collection
    Set sections = New Collection
    Set paragraphs = New Collection
    Set performThru = New Collection

    Dim rxSec As Object, rxPara As Object, rxThru As Object
    Set rxSec = CreateObject("VBScript.RegExp")
    rxSec.Pattern = "^([A-Z0-9][A-Z0-9_-]+)\s+SECTION$"
    rxSec.IgnoreCase = False
    Set rxPara = CreateObject("VBScript.RegExp")
    rxPara.Pattern = "^([A-Z0-9][A-Z0-9_-]+)$"
    rxPara.IgnoreCase = False
    Set rxThru = CreateObject("VBScript.RegExp")
    rxThru.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9_-]+)\s+THR(U|OUGH)\s+([A-Z0-9][A-Z0-9_-]+)"
    rxThru.IgnoreCase = False

    Dim entry As OrderedDict
    Dim line As OrderedDict, txt As String, m As Object, nm As String
    For Each line In lines
        txt = Convert_StripTrailingPeriod(line.Item("Text"))

        Set m = rxSec.Execute(txt)
        If m.Count > 0 Then
            Set entry = New OrderedDict
            entry.Add "name", UCase$(m.Item(0).SubMatches(0))
            entry.Add "line", line.Item("Number")
            sections.Add entry
            GoTo NextStructLine
        End If

        Set m = rxPara.Execute(txt)
        If m.Count > 0 Then
            nm = m.Item(0).SubMatches(0)
            If Test_ParagraphName(nm) Then
                Set entry = New OrderedDict
                entry.Add "name", UCase$(nm)
                entry.Add "line", line.Item("Number")
                paragraphs.Add entry
            End If
        End If

        Set m = rxThru.Execute(txt)
        If m.Count > 0 Then
            Set entry = New OrderedDict
            entry.Add "from", UCase$(m.Item(0).SubMatches(0))
            entry.Add "to", UCase$(m.Item(0).SubMatches(2))
            entry.Add "line", line.Item("Number")
            performThru.Add entry
        End If
NextStructLine:
    Next line

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "sections", sections
    result.Add "paragraphs", paragraphs
    result.Add "performThruRanges", performThru
    Set Get_ProgramStructure = result
End Function

'==============================================================================
' AST construction
'==============================================================================

Public Function Get_CobolAction(ByVal txt As String, ByVal lineNumber As Long) As OrderedDict
    Static rx As Object, rxSpaces As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "^(PERFORM|CALL|GO\s+TO|MOVE|COMPUTE|READ|WRITE|REWRITE|DELETE)\b\s*(.*)$"
        rx.IgnoreCase = False
        Set rxSpaces = CreateObject("VBScript.RegExp")
        rxSpaces.Pattern = "\s+"
        rxSpaces.Global = True
    End If
    Dim m As Object
    Set m = rx.Execute(txt)
    If m.Count = 0 Then
        Set Get_CobolAction = Nothing
        Exit Function
    End If
    Dim verb As String, rest As String, label As String
    verb = rxSpaces.Replace(UCase$(m.Item(0).SubMatches(0)), " ")
    rest = m.Item(0).SubMatches(1)
    label = Convert_CollapseSpaces(verb & " " & rest)

    Dim node As OrderedDict
    Set node = New OrderedDict
    node.Add "id", "action-" & lineNumber
    node.Add "type", "action"
    node.Add "label", label
    node.Add "startLine", lineNumber
    node.Add "endLine", lineNumber
    Set Get_CobolAction = node
End Function

' Build AST. Returns a Collection of root-level nodes.
Public Function Get_CobolNodes(ByVal lines As Collection) As Collection
    Dim root As OrderedDict
    Set root = New OrderedDict
    root.Add "type", "root"
    root.Add "children", New Collection

    Dim stack As Collection
    Set stack = New Collection

    Dim rootFrame As OrderedDict
    Set rootFrame = New OrderedDict
    rootFrame.Add "type", "root"
    rootFrame.Add "node", root
    rootFrame.Add "branch", "children"
    stack.Add rootFrame

    Dim rxEndIf As Object, rxEndEv As Object, rxEndSrch As Object
    Dim rxElse As Object, rxAtEnd As Object
    Dim rxIf As Object, rxEval As Object, rxSrch As Object, rxWhen As Object
    Set rxEndIf = CreateObject("VBScript.RegExp"):   rxEndIf.Pattern = "^END-IF\b"
    Set rxEndEv = CreateObject("VBScript.RegExp"):   rxEndEv.Pattern = "^END-EVALUATE\b"
    Set rxEndSrch = CreateObject("VBScript.RegExp"): rxEndSrch.Pattern = "^END-SEARCH\b"
    Set rxElse = CreateObject("VBScript.RegExp"):    rxElse.Pattern = "^ELSE\b"
    Set rxAtEnd = CreateObject("VBScript.RegExp"):   rxAtEnd.Pattern = "^AT\s+END\b"
    Set rxIf = CreateObject("VBScript.RegExp"):      rxIf.Pattern = "^IF\s+(.+)$"
    Set rxEval = CreateObject("VBScript.RegExp"):    rxEval.Pattern = "^EVALUATE\s+(.+)$"
    Set rxSrch = CreateObject("VBScript.RegExp"):    rxSrch.Pattern = "^SEARCH\s+(ALL\s+)?(.+)$"
    Set rxWhen = CreateObject("VBScript.RegExp"):    rxWhen.Pattern = "^WHEN\s+(.+)$"

    Dim line As OrderedDict, txt As String, m As Object, frame As OrderedDict, lineNum As Long
    Dim ifNode As OrderedDict, ifFrame As OrderedDict
    Dim evNode As OrderedDict, evFrame As OrderedDict
    Dim srNode As OrderedDict, srFrame As OrderedDict
    Dim wNode As OrderedDict, action As OrderedDict
    For Each line In lines
        lineNum = CLng(line.Item("Number"))
        txt = Convert_StripTrailingPeriod(line.Item("Text"))

        If rxEndIf.Test(txt) Then
            Close_Frame stack, "if", lineNum
            GoTo NextLine
        End If
        If rxEndEv.Test(txt) Then
            Close_Frame stack, "evaluate", lineNum
            GoTo NextLine
        End If
        If rxEndSrch.Test(txt) Then
            Close_Frame stack, "search", lineNum
            GoTo NextLine
        End If

        If rxElse.Test(txt) Then
            Set frame = Find_OpenFrame(stack, "if")
            If Not frame Is Nothing Then
                frame.Item("node").Add "elseLine", lineNum
                frame.Add "branch", "elseChildren"
            End If
            GoTo NextLine
        End If

        If rxAtEnd.Test(txt) Then
            Set frame = Find_OpenFrame(stack, "search")
            If Not frame Is Nothing Then
                frame.Item("node").Add "atEndLine", lineNum
                frame.Add "branch", "atEndChildren"
            End If
            GoTo NextLine
        End If

        Set m = rxIf.Execute(txt)
        If m.Count > 0 Then
            Set ifNode = New OrderedDict
            ifNode.Add "id", "if-" & lineNum
            ifNode.Add "type", "if"
            ifNode.Add "condition", Convert_NormalizeCondition(m.Item(0).SubMatches(0))
            ifNode.Add "startLine", lineNum
            ifNode.Add "endLine", lineNum
            ifNode.Add "elseLine", Null
            ifNode.Add "thenChildren", New Collection
            ifNode.Add "elseChildren", New Collection
            Add_ChildNode stack, ifNode
            Set ifFrame = New OrderedDict
            ifFrame.Add "type", "if"
            ifFrame.Add "node", ifNode
            ifFrame.Add "branch", "thenChildren"
            ifFrame.Add "currentCase", Nothing
            stack.Add ifFrame
            GoTo NextLine
        End If

        Set m = rxEval.Execute(txt)
        If m.Count > 0 Then
            Set evNode = New OrderedDict
            evNode.Add "id", "evaluate-" & lineNum
            evNode.Add "type", "evaluate"
            evNode.Add "expression", Convert_NormalizeCondition(m.Item(0).SubMatches(0))
            evNode.Add "startLine", lineNum
            evNode.Add "endLine", lineNum
            evNode.Add "cases", New Collection
            Add_ChildNode stack, evNode
            Set evFrame = New OrderedDict
            evFrame.Add "type", "evaluate"
            evFrame.Add "node", evNode
            evFrame.Add "branch", "cases"
            evFrame.Add "currentCase", Nothing
            stack.Add evFrame
            GoTo NextLine
        End If

        Set m = rxSrch.Execute(txt)
        If m.Count > 0 Then
            Set srNode = New OrderedDict
            srNode.Add "id", "search-" & lineNum
            srNode.Add "type", "search"
            srNode.Add "isAll", (Len(m.Item(0).SubMatches(0)) > 0)
            srNode.Add "tableExpr", Convert_NormalizeCondition(m.Item(0).SubMatches(1))
            srNode.Add "startLine", lineNum
            srNode.Add "endLine", lineNum
            srNode.Add "atEndLine", Null
            srNode.Add "atEndChildren", New Collection
            srNode.Add "cases", New Collection
            Add_ChildNode stack, srNode
            Set srFrame = New OrderedDict
            srFrame.Add "type", "search"
            srFrame.Add "node", srNode
            srFrame.Add "branch", "cases"
            srFrame.Add "currentCase", Nothing
            stack.Add srFrame
            GoTo NextLine
        End If

        Set m = rxWhen.Execute(txt)
        If m.Count > 0 Then
            Set frame = Find_OpenCaseFrame(stack)
            If Not frame Is Nothing Then
                Set wNode = New OrderedDict
                wNode.Add "id", "when-" & lineNum
                wNode.Add "type", "when"
                wNode.Add "condition", Convert_NormalizeCondition(m.Item(0).SubMatches(0))
                wNode.Add "startLine", lineNum
                wNode.Add "endLine", lineNum
                wNode.Add "children", New Collection
                frame.Item("node").Item("cases").Add wNode
                frame.Add "currentCase", wNode
                If frame.Item("type") = "search" Then frame.Add "branch", "cases"
            End If
            GoTo NextLine
        End If

        Set action = Get_CobolAction(txt, lineNum)
        If Not action Is Nothing Then Add_ChildNode stack, action

NextLine:
    Next line

    mUnclosedFrames = stack.Count - 1
    Set Get_CobolNodes = root.Item("children")
End Function

Private Sub Add_ChildNode(ByVal stack As Collection, ByVal item As OrderedDict)
    Dim frame As OrderedDict, node As OrderedDict, list As Collection, ftype As String, cur As Object
    Set frame = stack.Item(stack.Count)
    Set node = frame.Item("node")
    ftype = frame.Item("type")

    If ftype = "if" Then
        Set list = node.Item(CStr(frame.Item("branch")))
    ElseIf ftype = "evaluate" Then
        Set cur = frame.Item("currentCase")
        If cur Is Nothing Then
            Set list = node.Item("cases")
        Else
            Set list = cur.Item("children")
        End If
    ElseIf ftype = "search" Then
        If frame.Item("branch") = "atEndChildren" Then
            Set list = node.Item("atEndChildren")
        Else
            Set cur = frame.Item("currentCase")
            If cur Is Nothing Then
                Set list = node.Item("cases")
            Else
                Set list = cur.Item("children")
            End If
        End If
    Else
        Set list = node.Item("children")
    End If
    list.Add item
End Sub

Private Function Find_OpenFrame(ByVal stack As Collection, ByVal typ As String) As OrderedDict
    Dim i As Long, frame As OrderedDict
    For i = stack.Count To 1 Step -1
        Set frame = stack.Item(i)
        If frame.Item("type") = typ Then
            Set Find_OpenFrame = frame
            Exit Function
        End If
    Next i
    Set Find_OpenFrame = Nothing
End Function

Private Function Find_OpenCaseFrame(ByVal stack As Collection) As OrderedDict
    Dim i As Long, frame As OrderedDict, t As String
    For i = stack.Count To 1 Step -1
        Set frame = stack.Item(i)
        t = frame.Item("type")
        If t = "evaluate" Or t = "search" Then
            Set Find_OpenCaseFrame = frame
            Exit Function
        End If
    Next i
    Set Find_OpenCaseFrame = Nothing
End Function

Private Sub Close_Frame(ByVal stack As Collection, ByVal typ As String, ByVal endLine As Long)
    Dim i As Long, frame As OrderedDict
    For i = stack.Count To 2 Step -1
        Set frame = stack.Item(i)
        stack.Remove i
        If frame.Item("type") = typ Then
            frame.Item("node").Add "endLine", endLine
            Exit Sub
        End If
    Next i
End Sub

Public Function Get_NodeCount(ByVal nodes As Collection, ByVal types As Variant) As Long
    Dim cnt As Long, n As OrderedDict, t As String, i As Long, found As Boolean
    For Each n In nodes
        t = n.Item("type")
        found = False
        For i = LBound(types) To UBound(types)
            If types(i) = t Then
                found = True
                Exit For
            End If
        Next i
        If found Then cnt = cnt + 1
        If n.Exists("thenChildren") Then cnt = cnt + Get_NodeCount(n.Item("thenChildren"), types)
        If n.Exists("elseChildren") Then cnt = cnt + Get_NodeCount(n.Item("elseChildren"), types)
        If n.Exists("children") Then cnt = cnt + Get_NodeCount(n.Item("children"), types)
        If n.Exists("cases") Then cnt = cnt + Get_NodeCount(n.Item("cases"), types)
        If n.Exists("atEndChildren") Then cnt = cnt + Get_NodeCount(n.Item("atEndChildren"), types)
    Next n
    Get_NodeCount = cnt
End Function

'==============================================================================
' Call relationships + branch coverage (Phase 3)
'==============================================================================

Public Function Get_CallRelationships(ByVal lines As Collection, ByVal structure As OrderedDict) As OrderedDict
    Dim owners As Collection, o As OrderedDict
    Set owners = New Collection

    Dim s As OrderedDict
    For Each s In structure.Item("sections")
        Set o = New OrderedDict
        o.Add "name", s.Item("name")
        o.Add "line", s.Item("line")
        owners.Add o
    Next s
    Dim p As OrderedDict
    For Each p In structure.Item("paragraphs")
        Set o = New OrderedDict
        o.Add "name", p.Item("name")
        o.Add "line", p.Item("line")
        owners.Add o
    Next p
    Set owners = SortByLine_(owners)

    Dim edges As Collection
    Set edges = New Collection

    Dim rxThru As Object, rxPerf As Object, rxCall As Object, rxGoto As Object
    Set rxThru = CreateObject("VBScript.RegExp")
    rxThru.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9_-]+)\s+THR(U|OUGH)\s+([A-Z0-9][A-Z0-9_-]+)"
    Set rxPerf = CreateObject("VBScript.RegExp")
    rxPerf.Pattern = "^PERFORM\s+([A-Z0-9][A-Z0-9_-]+)"
    Set rxCall = CreateObject("VBScript.RegExp")
    rxCall.Pattern = "^CALL\s+'?([A-Z0-9][A-Za-z0-9_-]*)'?"
    Set rxGoto = CreateObject("VBScript.RegExp")
    rxGoto.Pattern = "^GO\s+TO\s+([A-Z0-9][A-Z0-9_-]+)"

    Dim line As OrderedDict, txt As String, ln As Long, owner As String, m As Object, edge As OrderedDict
    For Each line In lines
        txt = line.Item("Text")
        ln = CLng(line.Item("Number"))

        owner = "(MAIN)"
        For Each o In owners
            If CLng(o.Item("line")) <= ln Then
                owner = CStr(o.Item("name"))
            Else
                Exit For
            End If
        Next o

        Set m = rxThru.Execute(txt)
        If m.Count > 0 Then
            Set edge = New OrderedDict
            edge.Add "from", owner
            edge.Add "to", UCase$(m.Item(0).SubMatches(0))
            edge.Add "kind", "perform-thru"
            edge.Add "line", ln
            edges.Add edge
        Else
            Set m = rxPerf.Execute(txt)
            If m.Count > 0 Then
                Set edge = New OrderedDict
                edge.Add "from", owner
                edge.Add "to", UCase$(m.Item(0).SubMatches(0))
                edge.Add "kind", "perform"
                edge.Add "line", ln
                edges.Add edge
            End If
        End If

        Set m = rxCall.Execute(txt)
        If m.Count > 0 Then
            Set edge = New OrderedDict
            edge.Add "from", owner
            edge.Add "to", UCase$(m.Item(0).SubMatches(0))
            edge.Add "kind", "call"
            edge.Add "line", ln
            edges.Add edge
        End If

        Set m = rxGoto.Execute(txt)
        If m.Count > 0 Then
            Set edge = New OrderedDict
            edge.Add "from", owner
            edge.Add "to", UCase$(m.Item(0).SubMatches(0))
            edge.Add "kind", "goto"
            edge.Add "line", ln
            edges.Add edge
        End If
    Next line

    Dim nodes As Collection
    Set nodes = New Collection
    For Each s In structure.Item("sections")
        Set o = New OrderedDict
        o.Add "name", s.Item("name")
        o.Add "kind", "section"
        o.Add "line", s.Item("line")
        nodes.Add o
    Next s
    For Each p In structure.Item("paragraphs")
        Set o = New OrderedDict
        o.Add "name", p.Item("name")
        o.Add "kind", "paragraph"
        o.Add "line", p.Item("line")
        nodes.Add o
    Next p

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "nodes", nodes
    result.Add "edges", edges
    Set Get_CallRelationships = result
End Function

Public Function Get_BranchEdges(ByVal nodes As Collection) As Collection
    Dim edges As Collection
    Set edges = New Collection

    Dim n As OrderedDict, t As String, hasOther As Boolean
    Dim w As OrderedDict, sub_ As Collection, v As Variant, e As OrderedDict
    For Each n In nodes
        t = n.Item("type")
        If t = "if" Then
            Set e = New OrderedDict
            e.Add "branchId", n.Item("id") & ":then"
            e.Add "type", "if"
            e.Add "label", n.Item("condition") & " (THEN)"
            e.Add "line", n.Item("startLine")
            edges.Add e

            Set e = New OrderedDict
            e.Add "branchId", n.Item("id") & ":else"
            e.Add "type", "if"
            e.Add "label", n.Item("condition") & " (ELSE)"
            e.Add "line", n.Item("startLine")
            edges.Add e

            Set sub_ = Get_BranchEdges(n.Item("thenChildren"))
            For Each v In sub_: edges.Add v: Next v
            Set sub_ = Get_BranchEdges(n.Item("elseChildren"))
            For Each v In sub_: edges.Add v: Next v
        ElseIf t = "evaluate" Then
            hasOther = False
            For Each w In n.Item("cases")
                Set e = New OrderedDict
                e.Add "branchId", w.Item("id")
                e.Add "type", "when"
                e.Add "label", "WHEN " & w.Item("condition")
                e.Add "line", w.Item("startLine")
                edges.Add e
                If UCase$(Trim$(CStr(w.Item("condition")))) = "OTHER" Then hasOther = True
                Set sub_ = Get_BranchEdges(w.Item("children"))
                For Each v In sub_: edges.Add v: Next v
            Next w
            If Not hasOther Then
                Set e = New OrderedDict
                e.Add "branchId", n.Item("id") & ":nomatch"
                e.Add "type", "evaluate"
                e.Add "label", n.Item("expression") & " (該当WHENなし)"
                e.Add "line", n.Item("startLine")
                edges.Add e
            End If
        ElseIf t = "search" Then
            Set e = New OrderedDict
            e.Add "branchId", n.Item("id") & ":atend"
            e.Add "type", "search"
            e.Add "label", "AT END (" & n.Item("tableExpr") & ")"
            e.Add "line", n.Item("startLine")
            edges.Add e
            For Each w In n.Item("cases")
                Set e = New OrderedDict
                e.Add "branchId", w.Item("id")
                e.Add "type", "when"
                e.Add "label", "WHEN " & w.Item("condition")
                e.Add "line", w.Item("startLine")
                edges.Add e
                Set sub_ = Get_BranchEdges(w.Item("children"))
                For Each v In sub_: edges.Add v: Next v
            Next w
            Set sub_ = Get_BranchEdges(n.Item("atEndChildren"))
            For Each v In sub_: edges.Add v: Next v
        End If
    Next n
    Set Get_BranchEdges = edges
End Function

Public Function Get_BranchCoverage(ByVal rootNodes As Collection, ByVal testCases As Collection) As OrderedDict
    Dim edges As Collection
    Set edges = Get_BranchEdges(rootNodes)

    Dim coveredSet As Object
    Set coveredSet = CreateObject("Scripting.Dictionary")
    Dim tc As OrderedDict, bid As Variant
    For Each tc In testCases
        For Each bid In tc.Item("branchIds")
            If Not coveredSet.Exists(CStr(bid)) Then coveredSet.Add CStr(bid), True
        Next bid
    Next tc

    Dim branches As Collection, branch As OrderedDict, e As OrderedDict
    Set branches = New Collection
    Dim coveredCount As Long
    For Each e In edges
        Dim isCov As Boolean
        isCov = coveredSet.Exists(CStr(e.Item("branchId")))
        If isCov Then coveredCount = coveredCount + 1
        Dim byTc As Collection, found As Boolean
        Set byTc = New Collection
        If isCov Then
            For Each tc In testCases
                found = False
                For Each bid In tc.Item("branchIds")
                    If CStr(bid) = CStr(e.Item("branchId")) Then
                        found = True
                        Exit For
                    End If
                Next bid
                If found Then byTc.Add tc.Item("testCaseId")
            Next tc
        End If
        Set branch = New OrderedDict
        branch.Add "branchId", e.Item("branchId")
        branch.Add "type", e.Item("type")
        branch.Add "label", e.Item("label")
        branch.Add "line", e.Item("line")
        branch.Add "covered", isCov
        branch.Add "byTestCases", byTc
        branches.Add branch
    Next e

    Dim total As Long, rate As Double
    total = edges.Count
    If total > 0 Then rate = Round(coveredCount / total, 4) Else rate = 0

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "totalBranches", total
    result.Add "coveredBranches", coveredCount
    result.Add "coverageRate", rate
    result.Add "branches", branches
    Set Get_BranchCoverage = result
End Function

' Insertion sort a Collection of OrderedDict by the "line" field (ascending).
Private Function SortByLine_(ByVal c As Collection) As Collection
    Dim out As Collection
    Set out = New Collection
    Dim item As OrderedDict, i As Long, inserted As Boolean
    For Each item In c
        inserted = False
        For i = 1 To out.Count
            If CLng(item.Item("line")) < CLng(out.Item(i).Item("line")) Then
                out.Add item, , i
                inserted = True
                Exit For
            End If
        Next i
        If Not inserted Then out.Add item
    Next item
    Set SortByLine_ = out
End Function

'==============================================================================
' Path enumeration + test cases (Phase 4)
'==============================================================================

Public Function Expand_NodeSequence(ByVal nodes As Collection, ByVal conditions As Collection, _
        ByVal actions As Collection, ByVal lines As Collection, ByVal branchIds As Collection) As Collection
    Dim states As Collection
    Set states = New Collection
    states.Add NewState_(conditions, actions, lines, branchIds)

    Dim node As OrderedDict, ntype As String, s As OrderedDict
    Dim ns As Collection, v As Variant
    For Each node In nodes
        If mPathTruncated Then Exit For
        ntype = node.Item("type")

        If ntype = "action" Then
            Set ns = New Collection
            For Each s In states
                ns.Add NewState_(s.Item("conditions"), _
                                 ExtendObj_(s.Item("actions"), node), _
                                 ExtendVal_(s.Item("lines"), node.Item("startLine")), _
                                 s.Item("branchIds"))
            Next s
            Set states = ns
        ElseIf ntype = "if" Then
            Set ns = New Collection
            For Each s In states
                Dim thenStates As Collection, elseStates As Collection
                Set thenStates = Expand_NodeSequence(node.Item("thenChildren"), _
                    ExtendVal_(s.Item("conditions"), node.Item("condition")), _
                    s.Item("actions"), _
                    ExtendVal_(s.Item("lines"), node.Item("startLine")), _
                    ExtendVal_(s.Item("branchIds"), node.Item("id") & ":then"))
                For Each v In thenStates: ns.Add v: Next v
                If node.Item("elseChildren").Count > 0 Then
                    Set elseStates = Expand_NodeSequence(node.Item("elseChildren"), _
                        ExtendVal_(s.Item("conditions"), Convert_InvertCondition(CStr(node.Item("condition")))), _
                        s.Item("actions"), _
                        ExtendVal_(s.Item("lines"), node.Item("startLine")), _
                        ExtendVal_(s.Item("branchIds"), node.Item("id") & ":else"))
                    For Each v In elseStates: ns.Add v: Next v
                End If
            Next s
            Set states = ns
            If states.Count > MAX_PATH_STATES Then mPathTruncated = True
        ElseIf ntype = "evaluate" Then
            Set ns = New Collection
            For Each s In states
                Dim w As OrderedDict, subStates As Collection
                For Each w In node.Item("cases")
                    Set subStates = Expand_NodeSequence(w.Item("children"), _
                        ExtendVal_(s.Item("conditions"), w.Item("condition")), _
                        s.Item("actions"), _
                        ExtendVal_(ExtendVal_(s.Item("lines"), node.Item("startLine")), w.Item("startLine")), _
                        ExtendVal_(s.Item("branchIds"), w.Item("id")))
                    For Each v In subStates: ns.Add v: Next v
                Next w
            Next s
            Set states = ns
            If states.Count > MAX_PATH_STATES Then mPathTruncated = True
        ElseIf ntype = "search" Then
            Set ns = New Collection
            For Each s In states
                Dim w2 As OrderedDict, subStates2 As Collection
                For Each w2 In node.Item("cases")
                    Set subStates2 = Expand_NodeSequence(w2.Item("children"), _
                        ExtendVal_(s.Item("conditions"), w2.Item("condition")), _
                        s.Item("actions"), _
                        ExtendVal_(ExtendVal_(s.Item("lines"), node.Item("startLine")), w2.Item("startLine")), _
                        ExtendVal_(s.Item("branchIds"), w2.Item("id")))
                    For Each v In subStates2: ns.Add v: Next v
                Next w2
                If node.Item("atEndChildren").Count > 0 Then
                    Dim atEndStates As Collection
                    Set atEndStates = Expand_NodeSequence(node.Item("atEndChildren"), _
                        ExtendVal_(s.Item("conditions"), "AT END (" & node.Item("tableExpr") & ")"), _
                        s.Item("actions"), _
                        ExtendVal_(s.Item("lines"), node.Item("startLine")), _
                        ExtendVal_(s.Item("branchIds"), node.Item("id") & ":atend"))
                    For Each v In atEndStates: ns.Add v: Next v
                End If
            Next s
            Set states = ns
            If states.Count > MAX_PATH_STATES Then mPathTruncated = True
        End If
    Next node

    Set Expand_NodeSequence = states
End Function

Private Function NewState_(ByVal conds As Collection, ByVal acts As Collection, _
                           ByVal lns As Collection, ByVal bids As Collection) As OrderedDict
    Dim s As OrderedDict
    Set s = New OrderedDict
    s.Add "conditions", conds
    s.Add "actions", acts
    s.Add "lines", lns
    s.Add "branchIds", bids
    Set NewState_ = s
End Function

' Returns a new Collection that is c + [value]. Original collection unchanged.
Private Function ExtendVal_(ByVal c As Collection, ByVal value As Variant) As Collection
    Dim out As Collection, v As Variant
    Set out = New Collection
    For Each v In c
        If IsObject(v) Then out.Add v Else out.Add v
    Next v
    If IsObject(value) Then out.Add value Else out.Add value
    Set ExtendVal_ = out
End Function

Private Function ExtendObj_(ByVal c As Collection, ByVal obj As Object) As Collection
    Dim out As Collection, v As Variant
    Set out = New Collection
    For Each v In c
        If IsObject(v) Then out.Add v Else out.Add v
    Next v
    out.Add obj
    Set ExtendObj_ = out
End Function

Public Function New_ScenarioName(ByVal conditions As Collection, ByVal idx As Long) As String
    If conditions.Count > 0 Then
        New_ScenarioName = "シナリオ" & idx & ": " & JoinCol_(conditions, " / ")
    Else
        New_ScenarioName = "シナリオ" & idx & ": 条件なし"
    End If
End Function

Public Function Get_Warnings(ByVal lines As Collection, ByVal rootNodes As Collection) As Collection
    Dim w As Collection
    Set w = New Collection

    Dim line As OrderedDict, sb As String, first As Boolean
    first = True
    For Each line In lines
        If Not first Then sb = sb & vbLf
        sb = sb & line.Item("Text")
        first = False
    Next line

    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = False

    rx.Pattern = "\bNEXT\s+SENTENCE\b"
    If rx.Test(sb) Then w.Add "NEXT SENTENCE が含まれています。分岐終端の確認が必要です。"

    rx.Pattern = "\bGO\s+TO\b"
    If rx.Test(sb) Then w.Add "GO TO が含まれています。パス解析結果を手動確認してください。"

    rx.Pattern = "\bPERFORM\b.+\bTHR(U|OUGH)\b"
    If rx.Test(sb) Then w.Add "PERFORM THRU が含まれています。段落範囲候補を構造情報として抽出しています。"

    If Get_NodeCount(rootNodes, Array("evaluate")) > 0 Then
        w.Add "EVALUATE 条件はWHEN単位で抽出しています。複合条件の排他性はレビュー対象です。"
    End If
    If Get_NodeCount(rootNodes, Array("search")) > 0 Then
        w.Add "SEARCH の WHEN は表データに依存します。テストデータ設計時に表内容の確認が必要です。"
    End If

    Dim rxIfStart As Object, rxAndOr As Object, compoundFound As Boolean
    Set rxIfStart = CreateObject("VBScript.RegExp"): rxIfStart.Pattern = "^(IF|EVALUATE|WHEN)\b"
    Set rxAndOr = CreateObject("VBScript.RegExp"): rxAndOr.Pattern = "\b(AND|OR)\b"
    For Each line In lines
        If rxIfStart.Test(line.Item("Text")) And rxAndOr.Test(line.Item("Text")) Then
            compoundFound = True
            Exit For
        End If
    Next line
    If compoundFound Then w.Add "複合条件 (AND/OR) は分解していません。条件網羅は手動で確認してください。"

    If mUnclosedFrames > 0 Then
        w.Add "END句が不足している可能性があります。未クローズのブロックが " & mUnclosedFrames & " 個あります。"
    End If
    If mPathTruncated Then
        w.Add "パス数が上限 (" & MAX_PATH_STATES & ") を超えたため、一部のパスは展開されていません。"
    End If

    Set Get_Warnings = w
End Function

Private Function JoinCol_(ByVal c As Collection, ByVal sep As String) As String
    Dim sb As String, first As Boolean, v As Variant
    first = True
    For Each v In c
        If Not first Then sb = sb & sep
        sb = sb & CStr(v)
        first = False
    Next v
    JoinCol_ = sb
End Function

Private Function SortUniqueLong_(ByVal c As Collection) As Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim v As Variant, n As Long
    For Each v In c
        n = CLng(v)
        If Not seen.Exists(n) Then seen.Add n, True
    Next v
    Dim cnt As Long: cnt = seen.Count
    If cnt = 0 Then
        Set SortUniqueLong_ = New Collection
        Exit Function
    End If
    Dim arr() As Long, i As Long, j As Long, tmp As Long
    ReDim arr(0 To cnt - 1)
    i = 0
    Dim k As Variant
    For Each k In seen.Keys
        arr(i) = CLng(k)
        i = i + 1
    Next k
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j) < arr(i) Then
                tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            End If
        Next j
    Next i
    Dim out As Collection
    Set out = New Collection
    For i = LBound(arr) To UBound(arr)
        out.Add arr(i)
    Next i
    Set SortUniqueLong_ = out
End Function

'==============================================================================
' Orchestrators
'==============================================================================

' Phase 1: minimal summary { programName, lines, prefix* }.
Public Function Analyze_Phase1(ByVal source As String) As OrderedDict
    Dim norm As OrderedDict
    Set norm = Get_NormalizedCobolLines(source, "")

    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim summary As OrderedDict
    Set summary = New OrderedDict
    summary.Add "programName", Get_ProgramName(lines)
    summary.Add "lines", lines.Count
    summary.Add "prefixDetected", norm.Item("PrefixDetected")
    summary.Add "prefixStyle", norm.Item("PrefixStyle")
    summary.Add "prefixRatio", norm.Item("PrefixRatio")

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "summary", summary
    Set Analyze_Phase1 = result
End Function

' Phase 2: + rootNodes + programStructure (no paths/coverage/calls yet).
Public Function Analyze_Phase2(ByVal source As String) As OrderedDict
    mUnclosedFrames = 0
    mPathTruncated = False

    Dim norm As OrderedDict
    Set norm = Get_NormalizedCobolLines(source, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim rootNodes As Collection
    Set rootNodes = Get_CobolNodes(lines)

    Dim structure As OrderedDict
    Set structure = Get_ProgramStructure(lines)

    Dim summary As OrderedDict
    Set summary = New OrderedDict
    summary.Add "programName", Get_ProgramName(lines)
    summary.Add "lines", lines.Count
    summary.Add "branchCount", Get_NodeCount(rootNodes, Array("if", "evaluate", "when", "search"))
    summary.Add "actionCount", Get_NodeCount(rootNodes, Array("action"))
    summary.Add "sectionCount", structure.Item("sections").Count
    summary.Add "paragraphCount", structure.Item("paragraphs").Count
    summary.Add "performThruCount", structure.Item("performThruRanges").Count
    summary.Add "prefixDetected", norm.Item("PrefixDetected")
    summary.Add "prefixStyle", norm.Item("PrefixStyle")
    summary.Add "prefixRatio", norm.Item("PrefixRatio")

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "summary", summary
    result.Add "rootNodes", rootNodes
    result.Add "programStructure", structure
    Set Analyze_Phase2 = result
End Function

' Phase 3: + callGraph + coverage (with empty test cases -> coverage=0 framework).
Public Function Analyze_Phase3(ByVal source As String) As OrderedDict
    mUnclosedFrames = 0
    mPathTruncated = False

    Dim norm As OrderedDict
    Set norm = Get_NormalizedCobolLines(source, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim rootNodes As Collection
    Set rootNodes = Get_CobolNodes(lines)
    Dim structure As OrderedDict
    Set structure = Get_ProgramStructure(lines)
    Dim emptyTc As Collection
    Set emptyTc = New Collection
    Dim coverage As OrderedDict
    Set coverage = Get_BranchCoverage(rootNodes, emptyTc)
    Dim callGraph As OrderedDict
    Set callGraph = Get_CallRelationships(lines, structure)

    Dim summary As OrderedDict
    Set summary = New OrderedDict
    summary.Add "programName", Get_ProgramName(lines)
    summary.Add "lines", lines.Count
    summary.Add "branchCount", Get_NodeCount(rootNodes, Array("if", "evaluate", "when", "search"))
    summary.Add "actionCount", Get_NodeCount(rootNodes, Array("action"))
    summary.Add "sectionCount", structure.Item("sections").Count
    summary.Add "paragraphCount", structure.Item("paragraphs").Count
    summary.Add "performThruCount", structure.Item("performThruRanges").Count
    summary.Add "prefixDetected", norm.Item("PrefixDetected")
    summary.Add "prefixStyle", norm.Item("PrefixStyle")
    summary.Add "prefixRatio", norm.Item("PrefixRatio")

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "summary", summary
    result.Add "rootNodes", rootNodes
    result.Add "coverage", coverage
    result.Add "callGraph", callGraph
    result.Add "programStructure", structure
    Set Analyze_Phase3 = result
End Function

' Phase 4 / 5: full pipeline. Output matches the JSON schema consumed by
' CobolLogicViewer.BuildCobolReport.
Public Function Analyze_Full(ByVal source As String, Optional ByVal forcePrefix As String = "", _
                             Optional ByVal encodingName As String = "utf-8") As OrderedDict
    mUnclosedFrames = 0
    mPathTruncated = False

    Dim norm As OrderedDict
    Set norm = Get_NormalizedCobolLines(source, forcePrefix)
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim rootNodes As Collection
    Set rootNodes = Get_CobolNodes(lines)

    Dim empties(1 To 4) As Collection
    Dim i As Long
    For i = 1 To 4: Set empties(i) = New Collection: Next i

    Dim rawStates As Collection
    Set rawStates = Expand_NodeSequence(rootNodes, empties(1), empties(2), empties(3), empties(4))

    Dim pathStates As Collection, s As OrderedDict
    Set pathStates = New Collection
    For Each s In rawStates
        If s.Item("actions").Count > 0 Then pathStates.Add s
    Next s

    Dim testCases As Collection, tc As OrderedDict
    Set testCases = New Collection
    Dim idx As Long: idx = 0
    For Each s In pathStates
        idx = idx + 1
        Dim srcLines As Collection
        Set srcLines = SortUniqueLong_(s.Item("lines"))
        Dim actionLabels As Collection, actionLines As Collection, a As Variant
        Set actionLabels = New Collection
        Set actionLines = New Collection
        For Each a In s.Item("actions")
            actionLabels.Add a.Item("label")
            actionLines.Add a.Item("startLine")
        Next a
        Dim condSummary As String
        condSummary = JoinCol_(s.Item("conditions"), " / ")

        Set tc = New OrderedDict
        tc.Add "id", "P" & Format$(idx, "000")
        tc.Add "testCaseId", "TC-" & Format$(idx, "000")
        tc.Add "scenarioName", New_ScenarioName(s.Item("conditions"), idx)
        tc.Add "conditions", s.Item("conditions")
        tc.Add "conditionSummary", condSummary
        tc.Add "actionLabels", actionLabels
        tc.Add "actionLines", actionLines
        tc.Add "expectedResult", JoinCol_(actionLabels, " / ")
        If Len(condSummary) > 0 Then
            tc.Add "inputData", condSummary
        Else
            tc.Add "inputData", "条件なし"
        End If
        tc.Add "expectedValue", JoinCol_(actionLabels, " / ")
        If idx = 1 Then tc.Add "priority", "高" Else tc.Add "priority", "中"
        tc.Add "sourceLines", srcLines
        tc.Add "sourceLineText", JoinCol_(srcLines, ", ")
        tc.Add "branchIds", s.Item("branchIds")
        testCases.Add tc
    Next s

    Dim structure As OrderedDict
    Set structure = Get_ProgramStructure(lines)
    Dim coverage As OrderedDict
    Set coverage = Get_BranchCoverage(rootNodes, testCases)
    Dim callGraph As OrderedDict
    Set callGraph = Get_CallRelationships(lines, structure)
    Dim warnings As Collection
    Set warnings = Get_Warnings(lines, rootNodes)

    Dim summary As OrderedDict
    Set summary = New OrderedDict
    summary.Add "programName", Get_ProgramName(lines)
    summary.Add "lines", lines.Count
    summary.Add "branchCount", Get_NodeCount(rootNodes, Array("if", "evaluate", "when", "search"))
    summary.Add "pathCount", pathStates.Count
    summary.Add "actionCount", Get_NodeCount(rootNodes, Array("action"))
    summary.Add "sectionCount", structure.Item("sections").Count
    summary.Add "paragraphCount", structure.Item("paragraphs").Count
    summary.Add "performThruCount", structure.Item("performThruRanges").Count
    summary.Add "prefixDetected", norm.Item("PrefixDetected")
    summary.Add "prefixStyle", norm.Item("PrefixStyle")
    summary.Add "prefixRatio", norm.Item("PrefixRatio")
    summary.Add "encoding", encodingName
    summary.Add "generatedAt", Format$(Now, "yyyy-mm-dd\Thh:nn:ss")
    summary.Add "pathTruncated", mPathTruncated

    Dim srcCol As Collection
    Set srcCol = New Collection
    Dim rawArr() As String, hi As Long
    rawArr = Split(Replace(source, Chr$(13), ""), Chr$(10))
    hi = UBound(rawArr)
    If hi >= 0 Then If rawArr(hi) = "" Then hi = hi - 1
    For i = 0 To hi
        srcCol.Add rawArr(i)
    Next i

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "summary", summary
    result.Add "rootNodes", rootNodes
    result.Add "testCases", testCases
    result.Add "coverage", coverage
    result.Add "callGraph", callGraph
    result.Add "programStructure", structure
    result.Add "warnings", warnings
    result.Add "source", srcCol
    Set Analyze_Full = result
End Function
