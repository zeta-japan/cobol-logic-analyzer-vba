Attribute VB_Name = "CobolStub"
' CobolStub - ver2.0 feature (3): generate test Driver + Dummy COBOL skeletons.
'
' BuildStubsSheet(cblPath) renders, on a "Driver_Dummy雛形" sheet:
'   - a test Driver for the analyzed program (declares its LINKAGE arguments,
'     sets them, DISPLAYs in, CALLs the target, DISPLAYs out)
'   - one Dummy sub for each external CALL the program makes (LINKAGE built
'     from the argument PICs found in the program's DATA DIVISION)
' Skeletons are starting points: argument values and return values are filled
' in by hand per test case.

Option Explicit

Public Sub BuildStubsSheet(ByVal cblPath As String)
    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim progName As String
    progName = CobolParser.Get_ProgramName(lines)
    Dim usingParams As Collection, calls As Collection, items As Collection, linkage As Collection
    Set usingParams = CobolCalls.Get_ProcedureUsing(lines)
    Set calls = CobolCalls.Get_ExternalCalls(lines)
    Set items = CobolData.Get_DataItems(lines)
    Set linkage = CollectLinkageLines_(lines)

    Dim picMap As OrderedDict
    Set picMap = BuildPicMap_(items)

    ' Build the output text (Collection of OrderedDict { text, kind }).
    Dim out As Collection
    Set out = New Collection
    AppendDriver_ out, progName, usingParams, linkage
    Emit_ out, "", "blank"

    Dim c As OrderedDict
    For Each c In calls
        AppendDummy_ out, c, picMap
        Emit_ out, "", "blank"
    Next c
    If calls.Count = 0 Then
        Emit_ out, "(外部 CALL が無いため Dummy サブはありません)", "note"
    End If

    Render_ out, progName
End Sub

Private Sub AppendDriver_(ByVal out As Collection, ByVal progName As String, _
                          ByVal usingParams As Collection, ByVal linkage As Collection)
    Emit_ out, "==== テスト用 Driver : " & progName & " を呼び出す ====", "head"
    If usingParams.Count = 0 And linkage.Count = 0 Then
        Emit_ out, "* この PGM は LINKAGE / USING が無いため、被呼出サブではありません。", "code"
        Emit_ out, "* (Driver は不要。下の Dummy のみ利用してください)", "code"
        Exit Sub
    End If

    Emit_ out, "      *--- テストケースは「テストケース候補」シート参照 ---", "code"
    Emit_ out, "       IDENTIFICATION DIVISION.", "code"
    Emit_ out, "       PROGRAM-ID. DRV-" & progName & ".", "code"
    Emit_ out, "       DATA DIVISION.", "code"
    Emit_ out, "       WORKING-STORAGE SECTION.", "code"
    Emit_ out, "      *--- 対象 " & progName & " の引数 (LINKAGE 由来) ---", "code"
    Dim v As Variant
    For Each v In linkage
        Emit_ out, "       " & CStr(v) & ".", "code"
    Next v
    Emit_ out, "       PROCEDURE DIVISION.", "code"
    Emit_ out, "       MAIN-RTN.", "code"
    Emit_ out, "      *--- (1) テストケースに応じて引数へ値を設定 ---", "code"
    For Each v In usingParams
        Emit_ out, "      *     MOVE <値> TO " & CStr(v), "code"
    Next v
    Emit_ out, "      *--- (2) 入力値を表示 ---", "code"
    Emit_ out, "           DISPLAY ""=== " & progName & " INPUT ===""", "code"
    For Each v In usingParams
        Emit_ out, "           DISPLAY """ & CStr(v) & " = "" " & CStr(v), "code"
    Next v
    Emit_ out, "      *--- (3) 対象を呼び出し ---", "code"
    Emit_ out, "           CALL """ & progName & """ USING " & JoinSp_(usingParams) & ".", "code"
    Emit_ out, "      *--- (4) 返却値を表示 ---", "code"
    Emit_ out, "           DISPLAY ""=== " & progName & " OUTPUT ===""", "code"
    For Each v In usingParams
        Emit_ out, "           DISPLAY """ & CStr(v) & " = "" " & CStr(v), "code"
    Next v
    Emit_ out, "           STOP RUN.", "code"
End Sub

Private Sub AppendDummy_(ByVal out As Collection, ByVal callRec As OrderedDict, ByVal picMap As OrderedDict)
    Dim subName As String
    subName = CStr(callRec.Item("program"))
    Dim args As Collection
    Set args = callRec.Item("args")

    Emit_ out, "==== Dummy サブ : " & subName & " (返却値はテストケースに応じて設定) ====", "head"
    Emit_ out, "       IDENTIFICATION DIVISION.", "code"
    Emit_ out, "       PROGRAM-ID. " & subName & ".", "code"
    Emit_ out, "       DATA DIVISION.", "code"
    Emit_ out, "       LINKAGE SECTION.", "code"
    Emit_ out, "      *--- 呼出元から渡される引数 (CALL USING より) ---", "code"
    Dim v As Variant, pic As String
    For Each v In args
        If picMap.Exists(CStr(v)) Then
            pic = CStr(picMap.Item(CStr(v)))
        Else
            pic = ""
        End If
        If pic <> "" Then
            Emit_ out, "       01  " & CStr(v) & "  PIC " & pic & ".", "code"
        Else
            Emit_ out, "       01  " & CStr(v) & "  PIC X.   *> 桁/型はサブ仕様に合わせて修正", "code"
        End If
    Next v
    If args.Count > 0 Then
        Emit_ out, "       PROCEDURE DIVISION USING " & JoinSp_(args) & ".", "code"
    Else
        Emit_ out, "       PROCEDURE DIVISION.", "code"
    End If
    Emit_ out, "       DUMMY-RTN.", "code"
    Emit_ out, "      *--- テストケースに応じて返却値を設定 ---", "code"
    Emit_ out, "      *     MOVE <返却値> TO <出力引数>", "code"
    Emit_ out, "           GOBACK.", "code"
End Sub

' Collect the LINKAGE SECTION item lines (normalized text), for re-declaration
' in the Driver's WORKING-STORAGE. Faithful to COPY / OCCURS / PREFIXING.
Private Function CollectLinkageLines_(ByVal lines As Collection) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim rxSec As Object
    Set rxSec = CreateObject("VBScript.RegExp")
    rxSec.Pattern = "^(WORKING-STORAGE|LINKAGE|FILE|LOCAL-STORAGE)\s+SECTION$"
    rxSec.IgnoreCase = False
    Dim rxDiv As Object
    Set rxDiv = CreateObject("VBScript.RegExp")
    rxDiv.Pattern = "\bDIVISION\b"
    rxDiv.IgnoreCase = False

    Dim line As OrderedDict, txt As String, m As Object, inLink As Boolean
    inLink = False
    For Each line In lines
        txt = CobolParser.Convert_StripTrailingPeriod(line.Item("Text"))
        Set m = rxSec.Execute(txt)
        If m.Count > 0 Then
            inLink = (m.Item(0).SubMatches(0) = "LINKAGE")
        ElseIf inLink Then
            If rxDiv.Test(txt) Then
                Exit For ' reached PROCEDURE DIVISION
            ElseIf txt <> "" Then
                result.Add txt
            End If
        End If
    Next line
    Set CollectLinkageLines_ = result
End Function

Private Function BuildPicMap_(ByVal items As Collection) As OrderedDict
    Dim map As OrderedDict
    Set map = New OrderedDict
    Dim it As OrderedDict
    For Each it In items
        If CStr(it.Item("pic")) <> "" And Not map.Exists(CStr(it.Item("name"))) Then
            map.Add CStr(it.Item("name")), it.Item("pic")
        End If
    Next it
    Set BuildPicMap_ = map
End Function

Private Function JoinSp_(ByVal c As Collection) As String
    Dim s As String, v As Variant, first As Boolean
    first = True
    For Each v In c
        If Not first Then s = s & " "
        s = s & CStr(v)
        first = False
    Next v
    JoinSp_ = s
End Function

Private Sub Emit_(ByVal out As Collection, ByVal text As String, ByVal kind As String)
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "text", text
    e.Add "kind", kind
    out.Add e
End Sub

Private Sub Render_(ByVal out As Collection, ByVal progName As String)
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet("Driver_Dummy雛形")

    Application.ScreenUpdating = False
    On Error GoTo Done_
    ws.Cells.Clear

    ws.Range("A1").value = "Driver / Dummy 雛形 : " & progName
    With ws.Range("A1").Font
        .Name = "Meiryo UI": .Size = 14: .Bold = True: .Color = RGB(38, 70, 83)
    End With
    ws.Range("A2").value = "※ 雛形です。引数値・返却値はテストケースに応じて手で設定してください。COPY句は本体未展開です。"
    With ws.Range("A2").Font
        .Name = "Meiryo UI": .Size = 9: .Color = RGB(120, 120, 120)
    End With

    Dim row As Long, e As OrderedDict, kind As String
    row = 4
    For Each e In out
        kind = CStr(e.Item("kind"))
        ws.Cells(row, 1).value = CStr(e.Item("text"))
        Select Case kind
            Case "head"
                With ws.Cells(row, 1)
                    .Font.Bold = True
                    .Font.Name = "Meiryo UI"
                    .Interior.Color = RGB(217, 225, 232)
                End With
            Case "note"
                ws.Cells(row, 1).Font.Color = RGB(120, 120, 120)
                ws.Cells(row, 1).Font.Name = "Meiryo UI"
            Case Else
                ws.Cells(row, 1).Font.Name = "MS Gothic"
                ws.Cells(row, 1).Font.Size = 10
        End Select
        row = row + 1
    Next e

    ws.Columns("A").ColumnWidth = 90
    ws.Range("A1").Select
Done_:
    Application.ScreenUpdating = True
End Sub
