Attribute VB_Name = "Test_Phase7"
' Test_Phase7 - ver2.0 feature (2): DATA DIVISION item / PIC extraction.
' Validates CobolData against sample_calls.cbl and ICASE2.cbl plus unit
' tests for the PIC parser. (Oracle confirmed via the PS port.)

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_PicType"
    TestRunner.Run_One "Test_PicLen"
    TestRunner.Run_One "Test_DataItems_SampleCalls"
    TestRunner.Run_One "Test_DataItems_ICASE2"
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

Private Function FindItem_(ByVal items As Collection, ByVal name As String) As OrderedDict
    Dim it As OrderedDict
    For Each it In items
        If CStr(it.Item("name")) = name Then
            Set FindItem_ = it
            Exit Function
        End If
    Next it
    Set FindItem_ = Nothing
End Function

Public Sub Test_PicType()
    TestRunner.Assert_Equal "alnum", CobolData.PicType_("X(10)"), "PicType X(10)=alnum"
    TestRunner.Assert_Equal "num", CobolData.PicType_("9(05)"), "PicType 9(05)=num"
    TestRunner.Assert_Equal "signed-num", CobolData.PicType_("S9(5)"), "PicType S9(5)=signed-num"
    TestRunner.Assert_Equal "decimal", CobolData.PicType_("9(5)V99"), "PicType 9(5)V99=decimal"
End Sub

Public Sub Test_PicLen()
    TestRunner.Assert_Equal CLng(10), CLng(CobolData.PicLen_("X(10)")), "PicLen X(10)=10"
    TestRunner.Assert_Equal CLng(5), CLng(CobolData.PicLen_("9(05)")), "PicLen 9(05)=5"
    TestRunner.Assert_Equal CLng(3), CLng(CobolData.PicLen_("XXX")), "PicLen XXX=3"
    TestRunner.Assert_Equal CLng(7), CLng(CobolData.PicLen_("9(5)V99")), "PicLen 9(5)V99=7"
    TestRunner.Assert_Equal CLng(80), CLng(CobolData.PicLen_("X(80)")), "PicLen X(80)=80"
End Sub

Public Sub Test_DataItems_SampleCalls()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\sample_calls.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "sample_calls.cbl missing: " & p
        Exit Sub
    End If
    Dim items As Collection
    Set items = CobolData.Get_DataItems(LoadLines_("sample_calls.cbl"))
    Dim a As OrderedDict
    Set a = FindItem_(items, "AABB210")
    TestRunner.Assert_True Not a Is Nothing, "AABB210 found"
    TestRunner.Assert_Equal "alnum", CStr(a.Item("picType")), "AABB210 type=alnum"
    TestRunner.Assert_Equal CLng(10), CLng(a.Item("picLen")), "AABB210 len=10"
    Set a = FindItem_(items, "AABB220")
    TestRunner.Assert_Equal "num", CStr(a.Item("picType")), "AABB220 type=num"
    TestRunner.Assert_Equal CLng(5), CLng(a.Item("picLen")), "AABB220 len=5"
    Set a = FindItem_(items, "TBL0005-REC")
    TestRunner.Assert_True Not a Is Nothing, "TBL0005-REC found (FILE section)"
    TestRunner.Assert_Equal "FILE", CStr(a.Item("section")), "TBL0005-REC section=FILE"
End Sub

Public Sub Test_DataItems_ICASE2()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE2.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE2.cbl missing: " & p
        Exit Sub
    End If
    Dim items As Collection
    Set items = CobolData.Get_DataItems(LoadLines_("ICASE2.cbl"))
    Dim a As OrderedDict
    Set a = FindItem_(items, "WK-MODE")
    TestRunner.Assert_Equal "alnum", CStr(a.Item("picType")), "WK-MODE type=alnum"
    TestRunner.Assert_Equal CLng(1), CLng(a.Item("picLen")), "WK-MODE len=1"
    Set a = FindItem_(items, "WK-VAL")
    TestRunner.Assert_Equal "num", CStr(a.Item("picType")), "WK-VAL type=num"
    TestRunner.Assert_Equal CLng(5), CLng(a.Item("picLen")), "WK-VAL len=5"
    Set a = FindItem_(items, "WK-ENT")
    TestRunner.Assert_True Not a Is Nothing, "WK-ENT found"
    TestRunner.Assert_Equal "10", CStr(a.Item("occurs")), "WK-ENT occurs=10"
    Set a = FindItem_(items, "WK-AREA")
    TestRunner.Assert_Equal "group", CStr(a.Item("picType")), "WK-AREA type=group"
End Sub
