Attribute VB_Name = "Test_Phase10"
' Test_Phase10 - ver3.0 P1: statement merging (OR/AND continuation, EXEC ..
' END-EXEC, STRING .. INTO, CALL .. USING operands) + new verbs (EXEC/STRING/
' INITIALIZE/ACCEPT) + comment capture, validated on the ICASE3 fixture
' (cursor OPEN/FETCH/CLOSE + date sub + business sub + ABEND section).
' Expected values confirmed by the ver3 PS oracle (all assertions passed).

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_Merge_ICASE3"
    TestRunner.Run_One "Test_MergeRules_Edge"
End Sub

' Edge cases from the adversarial review of the merge/verb changes.
Public Sub Test_MergeRules_Edge()
    ' hyphenated paragraph names must NOT match as verbs
    TestRunner.Assert_True CobolParser.Get_CobolAction("INITIALIZE-RTN", 1) Is Nothing, _
        "INITIALIZE-RTN (paragraph) is not an action"
    TestRunner.Assert_True Not CobolParser.Get_CobolAction("INITIALIZE WK-AREA", 1) Is Nothing, _
        "INITIALIZE verb still an action"
    ' executable EXEC whose text merely contains DECLARE-like identifier stays an action
    TestRunner.Assert_True Not CobolParser.Get_CobolAction("EXEC ADABAS UPDATE WK-DECLARED-FLG END-EXEC", 1) Is Nothing, _
        "identifier containing DECLARE substring does not suppress EXEC action"
    TestRunner.Assert_True CobolParser.Get_CobolAction("EXEC ADABAS BEGIN DECLARE SECTION END-EXEC", 1) Is Nothing, _
        "DECLARE block still suppressed"

    ' leading-operator continuation style ("IF A = 1" / "OR B = 2")
    Dim c As Collection, m As Collection
    Set c = New Collection
    AddLine_ c, 1, "IF A = 1"
    AddLine_ c, 2, "OR B = 2"
    AddLine_ c, 3, "THEN"
    Set m = CobolParser.Merge_ContinuationLines(c)
    TestRunner.Assert_Equal CLng(2), CLng(m.Count), "leading-OR line merged into the IF"
    TestRunner.Assert_True InStr(CStr(m(1).Item("Text")), "OR B = 2") > 0, "merged condition keeps the OR term"

    ' an unterminated EXEC must stop at a section boundary (not eat the file)
    Set c = New Collection
    AddLine_ c, 1, "EXEC ADABAS"
    AddLine_ c, 2, "OPEN K77"
    AddLine_ c, 3, "NEXT-PROC SECTION."
    Set m = CobolParser.Merge_ContinuationLines(c)
    TestRunner.Assert_Equal CLng(2), CLng(m.Count), "unterminated EXEC stops at SECTION boundary"

    ' CALL with USING as the last token of the line still absorbs operands
    Set c = New Collection
    AddLine_ c, 1, "CALL 'SUBX' USING"
    AddLine_ c, 2, "WK-PARAM1"
    AddLine_ c, 3, "WK-PARAM2."
    AddLine_ c, 4, "MOVE A TO B"
    Set m = CobolParser.Merge_ContinuationLines(c)
    TestRunner.Assert_Equal CLng(2), CLng(m.Count), "USING-at-line-end CALL merged to the period"

    ' scope terminators are never absorbed as CALL operands
    Set c = New Collection
    AddLine_ c, 1, "CALL 'SUBX' USING WK-PARAM1"
    AddLine_ c, 2, "END-STRING."
    Set m = CobolParser.Merge_ContinuationLines(c)
    TestRunner.Assert_Equal CLng(2), CLng(m.Count), "END-STRING not absorbed as CALL operand"
End Sub

Private Sub AddLine_(ByVal c As Collection, ByVal n As Long, ByVal t As String)
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "Number", n
    e.Add "Raw", t
    e.Add "Text", t
    c.Add e
End Sub

Public Sub Test_Merge_ICASE3()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE3.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE3.cbl missing: " & p
        Exit Sub
    End If
    Dim src As String
    src = CobolEncoding.ReadCobolSource(p, "auto")
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    ' comment capture
    TestRunner.Assert_True norm.Exists("Comments"), "Comments captured in normalize result"
    TestRunner.Assert_True norm.Item("Comments").Count > 10, "ICASE3 comment count > 10"

    ' multi-line statements merged into one logical line each
    Dim entry As OrderedDict, mergedIf As String, mergedCall As String, mergedString As String
    For Each entry In lines
        If InStr(CStr(entry.Item("Text")), "IF DT2-MONTH") = 1 Then mergedIf = CStr(entry.Item("Text"))
        If InStr(CStr(entry.Item("Text")), "CALL 'ABSUB'") = 1 Then mergedCall = CStr(entry.Item("Text"))
        If InStr(CStr(entry.Item("Text")), "STRING ") = 1 Then mergedString = CStr(entry.Item("Text"))
    Next entry
    TestRunner.Assert_True InStr(mergedIf, "DT2-MONTH = 02") > 0 And InStr(mergedIf, "DT2-MONTH = 03") > 0, _
        "multi-line IF (OR continuation) merged"
    TestRunner.Assert_True InStr(mergedCall, "DT2-PARAM") > 0, "multi-line CALL USING merged to the period"
    TestRunner.Assert_True InStr(mergedString, "INTO IC3-PA200") > 0, "multi-line STRING merged through INTO"

    ' AST: 7 IF branches; EXEC actions only for executable blocks
    Dim root As Collection
    Set root = CobolParser.Get_CobolNodes(lines)
    Dim ifCount As Long, execCount As Long, stringCount As Long, declLeak As Long
    CountNodes_ root, ifCount, execCount, stringCount, declLeak
    TestRunner.Assert_Equal CLng(7), ifCount, "ICASE3 IF count = 7"
    TestRunner.Assert_Equal CLng(3), execCount, "EXEC actions = 3 (OPEN/FETCH/CLOSE; DECLARE blocks excluded)"
    TestRunner.Assert_Equal CLng(1), stringCount, "STRING action = 1"
    TestRunner.Assert_Equal CLng(0), declLeak, "no DECLARE block leaked as action"

    ' merged statement keeps the FIRST line's number
    Dim n As OrderedDict
    Set n = FindIf_(root, "DT2-MONTH")
    TestRunner.Assert_True Not n Is Nothing, "date IF node found"
    If Not n Is Nothing Then
        TestRunner.Assert_Equal CLng(117), CLng(n.Item("startLine")), "merged IF keeps first line number (117)"
    End If
End Sub

Private Sub CountNodes_(ByVal nodes As Collection, ByRef ifC As Long, ByRef exC As Long, ByRef stC As Long, ByRef dl As Long)
    Dim n As OrderedDict, t As String, lbl As String
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "if" Then
            ifC = ifC + 1
            CountNodes_ n.Item("thenChildren"), ifC, exC, stC, dl
            CountNodes_ n.Item("elseChildren"), ifC, exC, stC, dl
        ElseIf t = "evaluate" Then
            CountNodes_ n.Item("cases"), ifC, exC, stC, dl
        ElseIf t = "search" Then
            CountNodes_ n.Item("atEndChildren"), ifC, exC, stC, dl
            CountNodes_ n.Item("cases"), ifC, exC, stC, dl
        ElseIf t = "when" Then
            CountNodes_ n.Item("children"), ifC, exC, stC, dl
        ElseIf t = "action" Then
            lbl = CStr(n.Item("label"))
            If InStr(lbl, "EXEC ") = 1 Then exC = exC + 1
            If InStr(lbl, "STRING ") = 1 Then stC = stC + 1
            If InStr(lbl, "DECLARE") > 0 Then dl = dl + 1
        End If
    Next n
End Sub

Private Function FindIf_(ByVal nodes As Collection, ByVal needle As String) As OrderedDict
    Dim n As OrderedDict, t As String, r As OrderedDict
    Set FindIf_ = Nothing
    For Each n In nodes
        t = CStr(n.Item("type"))
        If t = "if" Then
            If InStr(CStr(n.Item("condition")), needle) > 0 Then
                Set FindIf_ = n
                Exit Function
            End If
            Set r = FindIf_(n.Item("thenChildren"), needle)
            If Not r Is Nothing Then Set FindIf_ = r: Exit Function
            Set r = FindIf_(n.Item("elseChildren"), needle)
            If Not r Is Nothing Then Set FindIf_ = r: Exit Function
        ElseIf t = "evaluate" Then
            Set r = FindIf_(n.Item("cases"), needle)
            If Not r Is Nothing Then Set FindIf_ = r: Exit Function
        ElseIf t = "search" Then
            Set r = FindIf_(n.Item("cases"), needle)
            If Not r Is Nothing Then Set FindIf_ = r: Exit Function
        ElseIf t = "when" Then
            Set r = FindIf_(n.Item("children"), needle)
            If Not r Is Nothing Then Set FindIf_ = r: Exit Function
        End If
    Next n
End Function
