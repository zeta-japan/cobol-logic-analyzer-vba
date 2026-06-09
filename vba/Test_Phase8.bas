Attribute VB_Name = "Test_Phase8"
' Test_Phase8 - ver2.0 feature (3): PROCEDURE DIVISION USING extraction
' (the basis of Driver generation). The Driver/Dummy text generation itself
' is reviewed visually on the "Driver_Dummystub" sheet.
' (Oracle confirmed via the PS port: ICASE1 -> ICA-PARAM, ICASE2 -> none.)

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_ProcUsing_ICASE1"
    TestRunner.Run_One "Test_ProcUsing_ICASE2"
    TestRunner.Run_One "Test_StubLines_ICASE1"
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

' Driver/Dummy text generation: BuildStubLines must produce real content
' (not an empty collection) for a program with LINKAGE + an external CALL.
' This isolates the "build" side from the sheet "render" side, and - because
' it references CobolStub - forces Run_All_Tests to COMPILE CobolStub.
Public Sub Test_StubLines_ICASE1()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE1.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE1.cbl missing: " & p
        Exit Sub
    End If

    Dim pn As String, out As Collection
    Set out = CobolStub.BuildStubLines(p, pn)

    TestRunner.Assert_True out.Count > 0, "ICASE1 stub lines built (count=" & out.Count & ")"
    TestRunner.Assert_Equal "ICASE1", pn, "ICASE1 stub progName=ICASE1"

    Dim foundDriver As Boolean, foundDummy As Boolean
    Dim e As OrderedDict, i As Long, t As String
    For i = 1 To out.Count
        Set e = out.Item(i)
        t = CStr(e.Item("text"))
        If InStr(t, "Driver") > 0 Then foundDriver = True
        If InStr(t, "ICASUB") > 0 Then foundDummy = True
    Next i
    TestRunner.Assert_True foundDriver, "ICASE1 stub contains a Driver section"
    TestRunner.Assert_True foundDummy, "ICASE1 stub contains an ICASUB Dummy section"
End Sub
