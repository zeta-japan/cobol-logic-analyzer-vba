Attribute VB_Name = "Test_Phase8"
' Test_Phase8 - ver2.0 feature (3): PROCEDURE DIVISION USING extraction
' (the basis of Driver generation). The Driver/Dummy text generation itself
' is reviewed visually on the "Driver_Dummystub" sheet.
' (Oracle confirmed via the PS port: ICASE1 -> ICA-PARAM, ICASE2 -> none.)

Option Explicit

Public Sub Run_All()
    TestRunner.Run_One "Test_ProcUsing_ICASE1"
    TestRunner.Run_One "Test_ProcUsing_ICASE2"
    TestRunner.Run_One "Test_DriverLines_ICASE3"
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

' ver3.0: per-case Driver generation. BuildDriverLines must produce one
' setup -> CALL -> DISPLAY block per NORMAL flow case, and no Dummy. Also
' exercises CobolIoView.BuildIoModel (the shared input/output derivation),
' forcing Run_All_Tests to COMPILE CobolStub + CobolIoView.
Public Sub Test_DriverLines_ICASE3()
    Dim p As String
    p = ThisWorkbook.path & "\samples\input\ICASE3.cbl"
    If Len(Dir(p)) = 0 Then
        TestRunner.Assert_True False, "ICASE3.cbl missing: " & p
        Exit Sub
    End If

    Dim src As String
    src = CobolEncoding.ReadCobolSource(p, "auto")
    Dim terms As Collection
    Set terms = New Collection
    terms.Add "S99-ABEND-PROC"
    Dim flow As OrderedDict
    Set flow = CobolFlow.Analyze_Flow(src, terms)

    Dim pn As String, out As Collection
    Set out = CobolStub.BuildDriverLines(flow, p, pn)

    TestRunner.Assert_True out.Count > 10, "ICASE3 driver lines built (count=" & out.Count & ")"
    TestRunner.Assert_Equal "ICASE3", pn, "driver progName=ICASE3"

    Dim callN As Long, dispN As Long, hasTC1 As Boolean, hasTC3 As Boolean, hasDummy As Boolean
    Dim e As OrderedDict, i As Long, t As String
    For i = 1 To out.Count
        Set e = out.Item(i)
        t = CStr(e.Item("text"))
        If InStr(t, "CALL 'ICASE3'") > 0 Then callN = callN + 1
        If InStr(t, "DISPLAY ") > 0 Then dispN = dispN + 1
        If InStr(t, "TC1") > 0 Then hasTC1 = True
        If InStr(t, "TC3") > 0 Then hasTC3 = True
        If InStr(t, "Dummy") > 0 Then hasDummy = True
    Next i
    TestRunner.Assert_Equal CLng(3), callN, "one CALL per normal case (3)"
    TestRunner.Assert_True dispN >= 3, "DISPLAY lines present per case"
    TestRunner.Assert_True hasTC1 And hasTC3, "case blocks TC1..TC3 present"
    TestRunner.Assert_True Not hasDummy, "no Dummy generation (retired)"

    ' the IO model behind the IO/expected-results sheet: 3 normal cases
    Dim model As Collection
    Set model = CobolIoView.BuildIoModel(flow, src)
    TestRunner.Assert_Equal CLng(3), CLng(model.Count), "IO model: 3 normal cases"
    Dim cm As OrderedDict
    Set cm = model(1)
    TestRunner.Assert_True cm.Item("outs").Count >= 3, "IO model TC1: linkage outputs derived"
    TestRunner.Assert_True cm.Item("dbPre").Count >= 1, "IO model TC1: DB prerequisites derived"
End Sub
