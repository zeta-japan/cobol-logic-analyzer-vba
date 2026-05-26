Attribute VB_Name = "Test_Phase8"
' Test_Phase8 - ver2.0 feature (3): PROCEDURE DIVISION USING extraction
' (the basis of Driver generation). The Driver/Dummy text generation itself
' is reviewed visually on the "Driver_Dummystub" sheet.
' (Oracle confirmed via the PS port: ICASE1 -> ICA-PARAM, ICASE2 -> none.)

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_ProcUsing_ICASE1"
    TestRunner.Run_One "Test_ProcUsing_ICASE2"
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

Public Sub Test_ProcUsing_ICASE1()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & p
        Exit Sub
    End If
    Dim u As Collection
    Set u = CobolCalls.Get_ProcedureUsing(LoadLines_("ICASE1.cbl"))
    TestRunner.Assert_Equal CLng(1), CLng(u.Count), "ICASE1 PROCEDURE USING count=1"
    TestRunner.Assert_Equal "ICA-PARAM", CStr(u.Item(1)), "ICASE1 USING param=ICA-PARAM"
End Sub

Public Sub Test_ProcUsing_ICASE2()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE2.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE2.cbl missing: " & p
        Exit Sub
    End If
    Dim u As Collection
    Set u = CobolCalls.Get_ProcedureUsing(LoadLines_("ICASE2.cbl"))
    TestRunner.Assert_Equal CLng(0), CLng(u.Count), "ICASE2 PROCEDURE USING count=0"
End Sub
