Attribute VB_Name = "Test_Phase6"
' Test_Phase6 - ver2.0 feature (1): external-call and data-access extraction.
' Validates CobolCalls against sample_calls.cbl and ICASE2.cbl.
' (Oracle confirmed independently via the PS port.)

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_Calls_SampleCalls"
    TestRunner.Run_One "Test_DataAccess_SampleCalls"
    TestRunner.Run_One "Test_Calls_ICASE2"
    TestRunner.Run_One "Test_DataAccess_ICASE2"
End Sub

Private Function LoadLines_(ByVal name As String) As Collection
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\" & name
    Dim src As String
    src = CobolEncoding.ReadCobolSource(p, "auto")
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Set LoadLines_ = norm.Item("Lines")
End Function

Private Function HasMode_(ByVal modes As Collection, ByVal mode As String) As Boolean
    Dim v As Variant
    For Each v In modes
        If v = mode Then
            HasMode_ = True
            Exit Function
        End If
    Next v
    HasMode_ = False
End Function

Public Sub Test_Calls_SampleCalls()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\sample_calls.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "sample_calls.cbl missing: " & p
        Exit Sub
    End If
    Dim calls As Collection
    Set calls = CobolCalls.Get_ExternalCalls(LoadLines_("sample_calls.cbl"))
    TestRunner.Assert_Equal CLng(3), CLng(calls.Count), "sample_calls: 3 external calls"
    Dim c As OrderedDict
    Set c = calls.Item(1)
    TestRunner.Assert_Equal "SUB001", CStr(c.Item("program")), "call1 = SUB001"
    TestRunner.Assert_Equal CLng(2), CLng(c.Item("args").Count), "SUB001: 2 args"
    TestRunner.Assert_Equal "AABB210", CStr(c.Item("args").Item(1)), "SUB001 arg1 = AABB210"
    TestRunner.Assert_Equal "AABB220", CStr(c.Item("args").Item(2)), "SUB001 arg2 = AABB220"
    Set c = calls.Item(3)
    TestRunner.Assert_Equal "SUB003", CStr(c.Item("program")), "call3 = SUB003"
    TestRunner.Assert_Equal CLng(1), CLng(c.Item("args").Count), "SUB003: 1 arg"
End Sub

Public Sub Test_DataAccess_SampleCalls()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\sample_calls.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "sample_calls.cbl missing: " & p
        Exit Sub
    End If
    Dim data As Collection
    Set data = CobolCalls.Get_DataAccess(LoadLines_("sample_calls.cbl"))
    TestRunner.Assert_Equal CLng(1), CLng(data.Count), "sample_calls: 1 data resource"
    Dim d As OrderedDict
    Set d = data.Item(1)
    TestRunner.Assert_Equal "TBL0005", CStr(d.Item("name")), "data = TBL0005 (record->file mapped)"
    TestRunner.Assert_True HasMode_(d.Item("modes"), "READ"), "TBL0005 has READ"
    TestRunner.Assert_True HasMode_(d.Item("modes"), "REWRITE"), "TBL0005 has REWRITE"
    TestRunner.Assert_True HasMode_(d.Item("modes"), "OPEN-INPUT"), "TBL0005 has OPEN-INPUT"
End Sub

Public Sub Test_Calls_ICASE2()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE2.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE2.cbl missing: " & p
        Exit Sub
    End If
    Dim calls As Collection
    Set calls = CobolCalls.Get_ExternalCalls(LoadLines_("ICASE2.cbl"))
    TestRunner.Assert_Equal CLng(1), CLng(calls.Count), "ICASE2: 1 external call"
    TestRunner.Assert_Equal "SUBCALC", CStr(calls.Item(1).Item("program")), "ICASE2 call = SUBCALC"
    TestRunner.Assert_Equal "WK-AREA", CStr(calls.Item(1).Item("args").Item(1)), "ICASE2 arg = WK-AREA"
End Sub

Public Sub Test_DataAccess_ICASE2()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE2.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE2.cbl missing: " & p
        Exit Sub
    End If
    Dim data As Collection
    Set data = CobolCalls.Get_DataAccess(LoadLines_("ICASE2.cbl"))
    TestRunner.Assert_Equal CLng(1), CLng(data.Count), "ICASE2: 1 data resource (IN-REC)"
    TestRunner.Assert_Equal "IN-REC", CStr(data.Item(1).Item("name")), "ICASE2 data = IN-REC"
End Sub
