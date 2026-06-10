Attribute VB_Name = "CobolStub"
' CobolStub - ver3.0 P5: per-test-case Driver skeleton on the Driver雛形
' sheet. Dummy generation was retired (team decision - subprograms are
' linked for real). One driver program runs every NORMAL case in sequence:
' per case it initializes the parameter, sets the linkage inputs
' (placeholders - the values live on the 入出力-想定結果 sheet), CALLs the
' target, and DISPLAYs the expected-output items for comparison.

Option Explicit

Private Const SHEET_DRV As String = "Driver雛形"

Public Sub BuildDriverSheet(ByVal flow As OrderedDict, ByVal cblPath As String)
    Dim progName As String
    Dim out As Collection
    On Error Resume Next
    Set out = BuildDriverLines(flow, cblPath, progName)
    On Error GoTo 0
    Render_ out, progName    ' Nothing-safe: Render_ writes a red diagnostic
End Sub

' Build the driver text as a Collection of OrderedDict { text, kind }.
' Public and sheet-free so Test_Phase8 can assert on the generated content.
Public Function BuildDriverLines(ByVal flow As OrderedDict, ByVal cblPath As String, _
                                 ByRef progNameOut As String) As Collection
    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim progName As String
    progName = CobolParser.Get_ProgramName(lines)
    progNameOut = progName
    Dim usingParams As Collection, linkage As Collection
    Set usingParams = CobolCalls.Get_ProcedureUsing(lines)
    Set linkage = CollectLinkageLines_(lines)

    Dim model As Collection
    Set model = CobolIoView.BuildIoModel(flow, src)

    Dim out As Collection
    Set out = New Collection
    Set BuildDriverLines = out

    Emit_ out, "==== テスト用 Driver : " & progName & " （正常系ケースを順次実行） ====", "head"
    If usingParams.Count = 0 And linkage.Count = 0 Then
        Emit_ out, "* この PGM は LINKAGE / USING が無いため、被呼出サブではありません。", "code"
        Emit_ out, "* (Driver は不要です)", "code"
        Exit Function
    End If
    If model.Count = 0 Then
        Emit_ out, "(正常系ケースが無いため Driver は生成されません)", "note"
        Exit Function
    End If

    Emit_ out, "      *--- 設定値・想定値は「入出力-想定結果」シート参照 ---", "code"
    Emit_ out, "       IDENTIFICATION DIVISION.", "code"
    Emit_ out, "       PROGRAM-ID. DRV-" & progName & ".", "code"
    Emit_ out, "      *  (PROGRAM-ID が 8 文字制限の環境では適宜短縮してください)", "code"
    Emit_ out, "       DATA DIVISION.", "code"
    Emit_ out, "       WORKING-STORAGE SECTION.", "code"
    Emit_ out, "      *--- 対象 " & progName & " の引数 (LINKAGE 由来) ---", "code"
    Dim v As Variant
    For Each v In linkage
        Emit_ out, "       " & CStr(v) & ".", "code"
    Next v
    Emit_ out, "       PROCEDURE DIVISION.", "code"
    Emit_ out, "       MAIN-RTN.", "code"

    Dim cm As OrderedDict
    For Each cm In model
        Emit_ out, "      *==============================================", "code"
        Emit_ out, "      *  " & CStr(cm.Item("id")) & " （正常系シナリオ" & CLng(cm.Item("kindSerial")) & "）", "code"
        Emit_ out, "      *  前提: " & PreSummary_(cm), "code"
        Emit_ out, "      *==============================================", "code"
        For Each v In usingParams
            Emit_ out, "           INITIALIZE " & CStr(v), "code"
        Next v
        Dim r As OrderedDict
        If cm.Item("lkIn").Count > 0 Then
            For Each r In cm.Item("lkIn")
                Emit_ out, "      *    " & CStr(cm.Item("id")) & " 入力: " & CStr(r.Item("Item")) & "  " & CStr(r.Item("Note")), "code"
                Emit_ out, "           MOVE SPACE TO " & CStr(r.Item("Item")), "code"
            Next r
        Else
            Emit_ out, "      *    （設定が必要な LINKAGE 入力はありません）", "code"
        End If
        If usingParams.Count > 0 Then
            Emit_ out, "           CALL '" & progName & "' USING " & JoinSp_(usingParams), "code"
        Else
            Emit_ out, "           CALL '" & progName & "'", "code"
        End If
        Emit_ out, "           DISPLAY '==== " & CStr(cm.Item("id")) & " RESULT ===='", "code"
        If cm.Item("outs").Count > 0 Then
            For Each r In cm.Item("outs")
                Emit_ out, "           DISPLAY '" & CStr(cm.Item("id")) & " " & CStr(r.Item("Item")) & " = ' " & CStr(r.Item("Item")), "code"
            Next r
        Else
            For Each v In usingParams
                Emit_ out, "           DISPLAY '" & CStr(cm.Item("id")) & " " & CStr(v) & " = ' " & CStr(v), "code"
            Next v
        End If
    Next cm

    Emit_ out, "           STOP RUN.", "code"
    Emit_ out, "", "blank"
    Emit_ out, "※ 異常系ケースは Driver 生成対象外です（機上確認推奨・テストケース候補シート参照）", "note"
End Function

Private Function PreSummary_(ByVal cm As OrderedDict) As String
    Dim s As String, v As Variant
    For Each v In cm.Item("dbPre")
        If Len(s) > 0 Then s = s & " ／ "
        s = s & CStr(v)
    Next v
    For Each v In cm.Item("subPre")
        If Len(s) > 0 Then s = s & " ／ "
        s = s & CStr(v)
    Next v
    If Len(s) = 0 Then s = "（特記なし）"
    PreSummary_ = s
End Function

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
    Set ws = JsonParser.EnsureSheet(SHEET_DRV)

    Application.ScreenUpdating = False
    ws.Cells.Clear

    ' Column A holds raw COBOL skeleton text. Header rows begin with "===="
    ' and Excel would treat a leading "=" as a FORMULA (runtime error 1004 on
    ' assignment). Forcing the column to Text makes every value store literally.
    ws.Columns(1).NumberFormat = "@"

    ws.Range("A1").Value = "Driver 雛形 : " & progName & " （正常系ケース毎に 設定 → CALL → DISPLAY）"
    With ws.Range("A1").Font
        .Name = "Meiryo UI": .Size = 14: .Bold = True: .Color = RGB(38, 70, 83)
    End With
    ws.Range("A2").Value = "※ 雛形です。MOVE の設定値は「入出力-想定結果」シートに合わせて手で記入してください。"
    With ws.Range("A2").Font
        .Name = "Meiryo UI": .Size = 9: .Color = RGB(120, 120, 120)
    End With

    Dim row As Long
    row = 4

    Dim n As Long
    If out Is Nothing Then n = -1 Else n = out.Count
    If n <= 0 Then
        ws.Cells(row, 1).Value = "(診断: 生成対象の行がありません。out=" & n & ")"
        ws.Cells(row, 1).Font.Color = RGB(192, 0, 0)
        GoTo Finish_
    End If

    Dim e As OrderedDict, kind As String, txt As String, i As Long
    For i = 1 To n
        kind = ""
        txt = ""
        On Error Resume Next
        Set e = out.Item(i)
        kind = CStr(e.Item("kind"))
        txt = CStr(e.Item("text"))
        ws.Cells(row, 1).Value = txt
        If Err.Number <> 0 Then
            kind = "note"
            ws.Cells(row, 1).Value = "(診断: 行 " & i & " でエラー #" & Err.Number & " " & Err.Description & ")"
            Err.Clear
        End If
        On Error GoTo 0

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
    Next i

Finish_:
    ws.Columns("A").ColumnWidth = 90
    On Error Resume Next
    Application.GoTo ws.Range("A1"), True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub
