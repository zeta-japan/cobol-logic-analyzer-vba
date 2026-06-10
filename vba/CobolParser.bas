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
Private mExpandCalls As Long
Private mExpandOps As Long   ' heartbeat counter (DoEvents) for big programs

Public Property Get UnclosedFrames() As Long
    UnclosedFrames = mUnclosedFrames
End Property

Public Property Get PathTruncated() As Boolean
    PathTruncated = mPathTruncated
End Property

Public Property Get ExpandCalls() As Long
    ExpandCalls = mExpandCalls
End Property

' Reset engine state. Useful when calling Expand_NodeSequence directly (e.g. a
' benchmark) outside of an Analyze_* orchestrator.
Public Sub ResetEngineState()
    mPathTruncated = False
    mUnclosedFrames = 0
    mExpandCalls = 0
End Sub

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
    Dim comments As Collection
    Set comments = New Collection

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
        If txt <> "" Then
            If rxComment.Test(txt) Then
                ' ver3.0: keep comments (stage labels for the test-case views)
                Set entry = New OrderedDict
                entry.Add "Number", idx
                entry.Add "Text", Trim$(Mid$(txt, 2))
                comments.Add entry
            Else
                Set entry = New OrderedDict
                entry.Add "Number", idx
                entry.Add "Raw", rl
                entry.Add "Text", txt
                lines.Add entry
            End If
        End If
    Next i

    ' ver3.0: stitch multi-line statements into one logical line
    Set lines = Merge_ContinuationLines(lines)

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "Lines", lines
    result.Add "PrefixDetected", detected
    result.Add "PrefixStyle", style
    result.Add "PrefixRatio", Round(ratio, 3)
    result.Add "Comments", comments
    Set Get_NormalizedCobolLines = result
End Function

' ver3.0 - merge multi-line statements into one logical line so the parser,
' path enumerator and call extractor all see complete statements:
'   1) a line ending with OR / AND continues on the next line (split conditions)
'   2) EXEC ... END-EXEC blocks become one statement
'   3) STRING continues until its INTO clause arrives
'   4) CALL ... USING continues over operand-only lines up to the period
' The merged statement keeps the FIRST line's number (hyperlinks/marking rely
' on it). Lines that trigger no rule pass through unchanged, so programs
' without continuations parse exactly as before.
Public Function Merge_ContinuationLines(ByVal lines As Collection) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim i As Long, cur As OrderedDict, txt As String, ntxt As String
    Dim take As Boolean, entry As OrderedDict
    i = 1
    Do While i <= lines.Count
        Set cur = lines(i)
        txt = CStr(cur.Item("Text"))
        Do While i + 1 <= lines.Count
            ntxt = CStr(lines(i + 1).Item("Text"))
            take = False
            If (Left$(txt, 5) = "EXEC " Or txt = "EXEC") And InStr(txt, "END-EXEC") = 0 Then
                ' guard: an unterminated EXEC must not swallow the whole file
                If Not Test_BlockBoundary(ntxt) Then take = True
            ElseIf Right$(txt, 3) = " OR" Or Right$(txt, 4) = " AND" Then
                take = True
            ElseIf Left$(ntxt, 3) = "OR " Or Left$(ntxt, 4) = "AND " Then
                ' leading-operator continuation style ("IF A = 1" / "OR B = 2")
                take = True
            ElseIf Left$(txt, 7) = "STRING " And InStr(txt, " INTO ") = 0 Then
                take = True
            ElseIf Left$(txt, 5) = "CALL " And InStr(txt & " ", " USING ") > 0 And Right$(txt, 1) <> "." Then
                If Test_OperandOnly(ntxt) Then take = True
            End If
            If Not take Then Exit Do
            txt = txt & " " & ntxt
            i = i + 1
        Loop
        Set entry = New OrderedDict
        entry.Add "Number", cur.Item("Number")
        entry.Add "Raw", cur.Item("Raw")
        entry.Add "Text", txt
        result.Add entry
        i = i + 1
    Loop
    Set Merge_ContinuationLines = result
End Function

' True if the line is a bare operand (identifier only, optional trailing
' period) and not a control keyword - i.e. a safe CALL-USING continuation.
Private Function Test_OperandOnly(ByVal t As String) As Boolean
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "^[A-Z0-9][A-Z0-9-]*\.?$"
        rx.IgnoreCase = False
    End If
    Test_OperandOnly = False
    If Not rx.Test(t) Then Exit Function
    Dim x As String
    x = t
    If Right$(x, 1) = "." Then x = Left$(x, Len(x) - 1)
    If Left$(x, 4) = "END-" Then Exit Function   ' any scope terminator
    If Not Test_ParagraphName(x) Then Exit Function
    Test_OperandOnly = True
End Function

' True for lines that begin a new section/division - used to stop a runaway
' merge when an EXEC block is missing its END-EXEC.
Private Function Test_BlockBoundary(ByVal t As String) As Boolean
    Test_BlockBoundary = False
    If InStr(t, " DIVISION") > 0 Then
        Test_BlockBoundary = True
        Exit Function
    End If
    Dim x As String
    x = t
    If Right$(x, 1) = "." Then x = Left$(x, Len(x) - 1)
    If Right$(x, 8) = " SECTION" Then
        ' only a real header (NAME SECTION, single token) is a boundary;
        ' "BEGIN DECLARE SECTION" inside EXEC .. END-EXEC is not.
        If InStr(Left$(x, Len(x) - 8), " ") = 0 Then Test_BlockBoundary = True
    End If
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
                     "END-READ", "END-CALL", "END-STRING", "END-COMPUTE", "END-ACCEPT", _
                     "END-EXEC", "END-WRITE", "END-REWRITE", "END-DELETE", "END-START", _
                     "EXIT", "CONTINUE", "GOBACK")
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
        ' (?!-) keeps hyphenated paragraph names (INITIALIZE-RTN etc.) from
        ' matching as verbs - a verb is always followed by a space or line end.
        rx.Pattern = "^(PERFORM|CALL|GO\s+TO|MOVE|COMPUTE|READ|WRITE|REWRITE|DELETE|STRING|INITIALIZE|ACCEPT|EXEC|GOBACK|STOP\s+RUN|EXIT\s+PROGRAM)\b(?!-)\s*(.*)$"
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
    If verb = "EXEC" Then
        ' DB declarations (BEGIN DECLARE SECTION / DECLARE ... CURSOR) are not
        ' executable statements - only OPEN/FETCH/CLOSE etc. become actions.
        ' Word-bounded so identifiers merely containing DECLARE do not match.
        Static rxDecl As Object
        If rxDecl Is Nothing Then
            Set rxDecl = CreateObject("VBScript.RegExp")
            rxDecl.Pattern = "\bDECLARE\b"
            rxDecl.IgnoreCase = False
        End If
        If rxDecl.Test(rest) Then
            Set Get_CobolAction = Nothing
            Exit Function
        End If
        rest = Trim$(Replace(rest, "END-EXEC", ""))
    End If
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

' Public compatibility wrapper: signature and result shape are unchanged
' (Collection seeds in, PathStates carrying ordered Collections out), but
' the enumeration itself runs on immutable cons lists (ConsList): extending
' a path is O(1) instead of a full Collection copy. The old copy-per-action
' behaviour made the expansion quadratic in program length and froze Excel
' on 1000+ line programs. Only the surviving states are materialized.
Public Function Expand_NodeSequence(ByVal nodes As Collection, ByVal conditions As Collection, _
        ByVal actions As Collection, ByVal lines As Collection, ByVal branchIds As Collection) As Collection
    mExpandOps = 0
    Dim states As Collection
    Set states = ExpandCons_(nodes, SeedCons_(conditions), SeedCons_(actions), _
                             SeedCons_(lines), SeedCons_(branchIds))
    Dim s As PathState
    For Each s In states
        MaterializeState_ s
    Next s
    Set Expand_NodeSequence = states
End Function

Private Function ExpandCons_(ByVal nodes As Collection, ByVal conditions As ConsList, _
        ByVal actions As ConsList, ByVal lines As ConsList, ByVal branchIds As ConsList) As Collection
    mExpandCalls = mExpandCalls + 1
    Dim states As Collection
    Set states = New Collection
    states.Add NewState_(conditions, actions, lines, branchIds)

    Dim node As OrderedDict, ntype As String, s As PathState
    Dim ns As Collection, v As Variant
    For Each node In nodes
        If mPathTruncated Then Exit For
        ntype = node.Item("type")

        If ntype = "action" Then
            Set ns = New Collection
            For Each s In states
                mExpandOps = mExpandOps + 1
                If (mExpandOps And 4095) = 0 Then DoEvents
                ns.Add NewState_(s.Conditions, _
                                 Cons_(s.Actions, node), _
                                 Cons_(s.Lines, node.Item("startLine")), _
                                 s.BranchIds)
            Next s
            Set states = ns
        ElseIf ntype = "if" Then
            Set ns = New Collection
            For Each s In states
                Dim thenStates As Collection, elseStates As Collection
                Set thenStates = ExpandCons_(node.Item("thenChildren"), _
                    Cons_(s.Conditions, node.Item("condition")), _
                    s.Actions, _
                    Cons_(s.Lines, node.Item("startLine")), _
                    Cons_(s.BranchIds, node.Item("id") & ":then"))
                For Each v In thenStates: ns.Add v: Next v
                If node.Item("elseChildren").Count > 0 Then
                    Set elseStates = ExpandCons_(node.Item("elseChildren"), _
                        Cons_(s.Conditions, Convert_InvertCondition(CStr(node.Item("condition")))), _
                        s.Actions, _
                        Cons_(s.Lines, node.Item("startLine")), _
                        Cons_(s.BranchIds, node.Item("id") & ":else"))
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
                    Set subStates = ExpandCons_(w.Item("children"), _
                        Cons_(s.Conditions, w.Item("condition")), _
                        s.Actions, _
                        Cons_(Cons_(s.Lines, node.Item("startLine")), w.Item("startLine")), _
                        Cons_(s.BranchIds, w.Item("id")))
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
                    Set subStates2 = ExpandCons_(w2.Item("children"), _
                        Cons_(s.Conditions, w2.Item("condition")), _
                        s.Actions, _
                        Cons_(Cons_(s.Lines, node.Item("startLine")), w2.Item("startLine")), _
                        Cons_(s.BranchIds, w2.Item("id")))
                    For Each v In subStates2: ns.Add v: Next v
                Next w2
                If node.Item("atEndChildren").Count > 0 Then
                    Dim atEndStates As Collection
                    Set atEndStates = ExpandCons_(node.Item("atEndChildren"), _
                        Cons_(s.Conditions, "AT END (" & node.Item("tableExpr") & ")"), _
                        s.Actions, _
                        Cons_(s.Lines, node.Item("startLine")), _
                        Cons_(s.BranchIds, node.Item("id") & ":atend"))
                    For Each v In atEndStates: ns.Add v: Next v
                End If
            Next s
            Set states = ns
            If states.Count > MAX_PATH_STATES Then mPathTruncated = True
        End If
    Next node

    Set ExpandCons_ = states
End Function

' During expansion the PathState fields hold ConsList heads (or Nothing);
' NewState_ just shares the immutable heads.
Private Function NewState_(ByVal conds As Object, ByVal acts As Object, _
                           ByVal lns As Object, ByVal bids As Object) As PathState
    Dim s As PathState
    Set s = New PathState
    Set s.Conditions = conds
    Set s.Actions = acts
    Set s.Lines = lns
    Set s.BranchIds = bids
    Set NewState_ = s
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

' Convert a (possibly empty) seed Collection into a cons list head.
Private Function SeedCons_(ByVal c As Collection) As ConsList
    Dim h As ConsList, v As Variant
    Set h = Nothing
    If c Is Nothing Then Exit Function
    For Each v In c
        Set h = Cons_(h, v)
    Next v
    Set SeedCons_ = h
End Function

' Materialize a finished state's cons heads back into ordered Collections
' (the public PathState contract consumed by the test-case builder/tests).
Private Sub MaterializeState_(ByVal s As PathState)
    Dim h As ConsList
    Set h = s.Conditions
    Set s.Conditions = ConsToColl_(h)
    Set h = s.Actions
    Set s.Actions = ConsToColl_(h)
    Set h = s.Lines
    Set s.Lines = ConsToColl_(h)
    Set h = s.BranchIds
    Set s.BranchIds = ConsToColl_(h)
End Sub

Private Function ConsToColl_(ByVal head As ConsList) As Collection
    Dim c As Collection
    Set c = New Collection
    Set ConsToColl_ = c
    If head Is Nothing Then Exit Function
    Dim n As Long, i As Long
    n = head.N
    Dim arr() As Variant
    ReDim arr(1 To n)
    Dim cur As ConsList
    Set cur = head
    For i = n To 1 Step -1
        If IsObject(cur.V) Then Set arr(i) = cur.V Else arr(i) = cur.V
        Set cur = cur.Prev
    Next i
    For i = 1 To n
        c.Add arr(i)
    Next i
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

    Dim pathStates As Collection, s As PathState
    Set pathStates = New Collection
    For Each s In rawStates
        If s.Actions.Count > 0 Then pathStates.Add s
    Next s

    Dim testCases As Collection, tc As OrderedDict
    Set testCases = New Collection
    Dim idx As Long: idx = 0
    For Each s In pathStates
        idx = idx + 1
        Dim srcLines As Collection
        Set srcLines = SortUniqueLong_(s.Lines)
        Dim actionLabels As Collection, actionLines As Collection, a As Variant
        Set actionLabels = New Collection
        Set actionLines = New Collection
        For Each a In s.Actions
            actionLabels.Add a.Item("label")
            actionLines.Add a.Item("startLine")
        Next a
        Dim condSummary As String
        condSummary = JoinCol_(s.Conditions, " / ")

        Set tc = New OrderedDict
        tc.Add "id", "P" & Format$(idx, "000")
        tc.Add "testCaseId", "TC-" & Format$(idx, "000")
        tc.Add "scenarioName", New_ScenarioName(s.Conditions, idx)
        tc.Add "conditions", s.Conditions
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
        tc.Add "branchIds", s.BranchIds
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
