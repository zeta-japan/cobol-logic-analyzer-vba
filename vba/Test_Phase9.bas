Attribute VB_Name = "Test_Phase9"
' Test_Phase9 - ver2.1: test-case consolidation + decision-path extraction
' (CobolTcMark.BuildTcGroups), the data behind the column-D tree marking.
'
' Oracle confirmed via the PS port against the real samples:
'   ICASE1 / ICASE2 : 96 paths -> 8 groups (consolidated by final Action).
'   Decision paths are the representative path's branch arms, truncated at the
'   final Action line. Reps are TC-001..TC-008 (fewest branches, lowest id).
' Running this also forces Run_All_Tests to COMPILE CobolTcMark.

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_TcGroups_ICASE1"
    TestRunner.Run_One "Test_TcGroups_ICASE2"
End Sub

Private Function Groups_(ByVal nm As String) As Collection
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\" & nm
    Dim src As String
    src = CobolEncoding.ReadCobolSource(p, "auto")
    Dim res As OrderedDict
    Set res = CobolParser.Analyze_Full(src, "", "utf-8")
    Set Groups_ = CobolTcMark.BuildTcGroups(res.Item("testCases"))
End Function

Private Function F_(ByVal groups As Collection, ByVal idx As Long, ByVal key As String) As Variant
    Dim g As OrderedDict
    Set g = groups.Item(idx)
    F_ = g.Item(key)
End Function

Public Sub Test_TcGroups_ICASE1()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & p
        Exit Sub
    End If
    Dim g As Collection
    Set g = Groups_("ICASE1.cbl")
    TestRunner.Assert_Equal CLng(8), CLng(g.Count), "ICASE1 -> 8 groups"

    ' all 8 final lines (grouping is independent of rep choice)
    Dim expLines As Variant, i As Long
    expLines = Array(112, 114, 117, 120, 123, 126, 129, 132)
    For i = 1 To 8
        TestRunner.Assert_Equal CLng(expLines(i - 1)), CLng(F_(g, i, "finalLine")), "ICASE1 G" & i & " finalLine"
        TestRunner.Assert_Equal CLng(12), CLng(F_(g, i, "memberCount")), "ICASE1 G" & i & " members=12"
    Next i

    ' representative decision path (G1 = deepest THEN chain, G8 = first ELSE)
    TestRunner.Assert_Equal "PERFORM MAIN-PREMIUM", CStr(F_(g, 1, "finalLabel")), "ICASE1 G1 label"
    TestRunner.Assert_Equal "TC-001", CStr(F_(g, 1, "repId")), "ICASE1 G1 rep"
    TestRunner.Assert_Equal "61:THEN 68:WHEN 80:WHEN 99:THEN 101:THEN 103:THEN 105:THEN 107:THEN 109:THEN 111:THEN", _
        CStr(F_(g, 1, "decisionPath")), "ICASE1 G1 decisionPath"
    TestRunner.Assert_Equal "PERFORM MAIN-NON-PREMIUM", CStr(F_(g, 8, "finalLabel")), "ICASE1 G8 label"
    TestRunner.Assert_Equal "61:THEN 68:WHEN 80:WHEN 99:ELSE", _
        CStr(F_(g, 8, "decisionPath")), "ICASE1 G8 decisionPath"
End Sub

Public Sub Test_TcGroups_ICASE2()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE2.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE2.cbl missing: " & p
        Exit Sub
    End If
    Dim g As Collection
    Set g = Groups_("ICASE2.cbl")
    TestRunner.Assert_Equal CLng(8), CLng(g.Count), "ICASE2 -> 8 groups"

    Dim expLines As Variant, i As Long
    expLines = Array(98, 100, 103, 106, 108, 111, 114, 117)
    For i = 1 To 8
        TestRunner.Assert_Equal CLng(expLines(i - 1)), CLng(F_(g, i, "finalLine")), "ICASE2 G" & i & " finalLine"
        TestRunner.Assert_Equal CLng(12), CLng(F_(g, i, "memberCount")), "ICASE2 G" & i & " members=12"
    Next i

    TestRunner.Assert_Equal "PERFORM MAIN-TOP", CStr(F_(g, 1, "finalLabel")), "ICASE2 G1 label"
    TestRunner.Assert_Equal "TC-001", CStr(F_(g, 1, "repId")), "ICASE2 G1 rep"
    TestRunner.Assert_Equal "50:THEN 57:WHEN 68:WHEN 86:THEN 88:THEN 90:THEN 93:WHEN 95:THEN 97:THEN", _
        CStr(F_(g, 1, "decisionPath")), "ICASE2 G1 decisionPath"
    TestRunner.Assert_Equal "GO TO ICASE2-MAIN-EXIT", CStr(F_(g, 8, "finalLabel")), "ICASE2 G8 label"
    TestRunner.Assert_Equal "50:THEN 57:WHEN 68:WHEN 86:ELSE", _
        CStr(F_(g, 8, "decisionPath")), "ICASE2 G8 decisionPath"
End Sub
