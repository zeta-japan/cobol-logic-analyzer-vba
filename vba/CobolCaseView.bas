Attribute VB_Name = "CobolCaseView"
' CobolCaseView - ver3.0 P3: render the generated test cases.
'   1) テストケース候補  - one vertical step-flow block per case (the format
'      the test team writes by hand: stage labels, external interactions with
'      their outcome arms, key assignments, and the termination form)
'   2) 分岐カバレッジ表  - rows = branch arms (検証Point), columns = cases,
'      ○ = the case's path takes that arm. An empty row = uncovered -> red.
' Input is CobolFlow.Analyze_Flow's result.

Option Explicit

' per-case render guard: a 1000+ line path carries thousands of step rows
Private Const MAX_CASE_STEPS As Long = 400

Private Const SHEET_CASES As String = "テストケース候補"
Private Const SHEET_MATRIX As String = "分岐カバレッジ表"

Public Sub BuildCaseSheets(ByVal flow As OrderedDict)
    On Error GoTo Done_
    Application.ScreenUpdating = False
    RenderCases_ flow
    RenderMatrix_ flow
Done_:
    Application.ScreenUpdating = True
End Sub

'======================================================================
' テストケース候補 (step-flow blocks)
'======================================================================
Private Sub RenderCases_(ByVal flow As OrderedDict)
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(SHEET_CASES)
    ws.Cells.Clear
    ws.Columns("A:B").NumberFormat = "@"

    ' title band: dark navy with white text (JP corporate sheet style)
    ws.Range("A1").Value = "テストケース候補（分岐網羅 C1 ／ ロジック階層(実行順展開)のケース標記・分岐カバレッジ表と対応）"
    With ws.Range("A1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
    End With
    ' which terminator sections were applied (auto/manual) - keep it visible
    ws.Cells(2, 1).Value = TermsNote_(flow)
    ws.Cells(2, 1).Font.Color = RGB(120, 120, 120)
    ws.Cells(2, 1).Font.Size = 9
    Dim row As Long
    row = 4
    If CBool(flow.Item("truncated")) Then
        ws.Cells(row, 1).Value = "※ パス数が上限を超えたため一部のパスを打ち切りました。カバレッジに漏れが出る可能性があります。"
        ws.Cells(row, 1).Font.Color = RGB(192, 0, 0)
        row = row + 2
    End If

    Dim c As OrderedDict
    For Each c In flow.Item("cases")
        row = RenderCaseBlock_(ws, c, row) + 1
    Next c

    ws.Columns("A").ColumnWidth = 100
End Sub

Private Function RenderCaseBlock_(ByVal ws As Worksheet, ByVal c As OrderedDict, ByVal startRow As Long) As Long
    Dim row As Long
    row = startRow

    ' case band: deep green (normal) / gray (out-of-scope), white text
    Dim kindJp As String, hdrColor As Long, isNormal As Boolean
    isNormal = (CStr(c.Item("kind")) = "normal")
    If isNormal Then
        kindJp = "正常系シナリオ" & CLng(c.Item("kindSerial"))
        hdrColor = RGB(55, 86, 35)
    Else
        kindJp = "異常系シナリオ" & CLng(c.Item("kindSerial"))
        hdrColor = RGB(89, 89, 89)
    End If

    ws.Cells(row, 1).Value = ChrW$(&H25A0) & " " & CStr(c.Item("id")) & "（" & kindJp & "）　　終了形態: " & TermJp_(c)
    With ws.Cells(row, 1)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = hdrColor
    End With
    row = row + 1

    If Not isNormal Then
        ws.Cells(row, 1).Value = "　※ 今回テスト対象外（機上確認推奨）"
        ws.Cells(row, 1).Font.Color = RGB(120, 120, 120)
        row = row + 1
    End If

    Dim e As OrderedDict, k As String, txt As String
    Dim stepN As Long, stepsCapped As Boolean
    For Each e In c.Item("events")
        k = CStr(e.Item("Kind"))
        ' count only events that render a row (cap = visible rows)
        Select Case k
            Case "enter", "exec", "call", "arm", "term"
                stepN = stepN + 1
            Case "assign"
                If CBool(e.Item("IsKey")) Then stepN = stepN + 1
        End Select
        If stepN > MAX_CASE_STEPS And k <> "term" Then
            If Not stepsCapped Then
                ws.Cells(row, 1).Value = "　…（以降の途中ステップは省略・終了行のみ表示）"
                ws.Cells(row, 1).Font.Color = RGB(120, 120, 120)
                row = row + 1
                stepsCapped = True
            End If
            GoTo NextEv
        End If
        txt = ""
        Select Case k
            Case "enter"
                If Len(CStr(e.Item("Label"))) > 0 Then
                    txt = "【" & CStr(e.Item("Label")) & "】 " & CStr(e.Item("Text"))
                Else
                    txt = "【" & CStr(e.Item("Text")) & "】"
                End If
                ws.Cells(row, 1).Value = txt
                With ws.Cells(row, 1)
                    .Font.Bold = True
                    .Font.Color = RGB(31, 78, 121)
                    .Interior.Color = RGB(220, 230, 241)
                End With
                row = row + 1
            Case "exec", "call"
                ws.Cells(row, 1).Value = "　　" & CStr(e.Item("Text"))
                row = row + 1
            Case "arm"
                ws.Cells(row, 1).Value = "　　　→ " & CStr(e.Item("Cond")) & " ： " & CStr(e.Item("Arm"))
                ws.Cells(row, 1).Font.Color = RGB(0, 112, 192)
                row = row + 1
            Case "assign"
                If CBool(e.Item("IsKey")) Then
                    ws.Cells(row, 1).Value = "　　" & CStr(e.Item("Dst")) & " ← " & CStr(e.Item("Src")) & "　（キー項目設定）"
                    row = row + 1
                End If
            Case "term"
                ws.Cells(row, 1).Value = "　終了　（" & TermJp_(c) & "）"
                With ws.Cells(row, 1)
                    .Font.Bold = True
                    .Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Borders(xlEdgeTop).Color = RGB(55, 86, 35)
                End With
                row = row + 1
        End Select
NextEv:
    Next e

    ' synthesized cases carry no term event of their own when built from a
    ' call-site snapshot whose Events end with the synthetic marker - ensure
    ' a closing line exists
    RenderCaseBlock_ = row
End Function

' What terminator sections were applied (auto-detected by ABEND naming or
' registered on the control sheet B24:B29) - shown so the user can verify.
' render the uncovered-arm reason code from the engine in Japanese
Private Function DiagJp_(ByVal flow As OrderedDict, ByVal token As String) As String
    DiagJp_ = ""
    If Not flow.Exists("armDiag") Then Exit Function
    Dim d As OrderedDict
    Set d = flow.Item("armDiag")
    If Not d.Exists(token) Then Exit Function
    Dim c As String
    c = CStr(d.Item(token))
    If Left$(c, 5) = "noctx" Then
        Dim nb As String, p1 As Long, p2 As Long, nSec As String, nCallers As String
        nb = Mid$(c, 7)
        Dim nRefs As String
        nRefs = ""
        p1 = InStr(nb, "|")
        If p1 > 0 Then
            nSec = Left$(nb, p1 - 1)
            nRefs = Mid$(nb, p1 + 1)
            p2 = InStr(nRefs, "|")
            If p2 > 0 Then
                nCallers = Left$(nRefs, p2 - 1)
                nRefs = Mid$(nRefs, p2 + 1)
            Else
                nCallers = nRefs
                nRefs = ""
            End If
        Else
            nSec = nb
            nCallers = ""
        End If
        If Len(nSec) = 0 Then
            DiagJp_ = "｜経路なし（PERFORM 未到達領域）"
        ElseIf Len(nCallers) > 0 Then
            DiagJp_ = "｜経路なし（" & nSec & "：呼出関係表では " & nCallers & " から呼出あり→解析ギャップの可能性、要連絡）"
        ElseIf Len(nRefs) > 0 Then
            DiagJp_ = "｜経路なし（" & nSec & "：参照行 " & Replace(nRefs, "~", " ／ ") & "）"
        Else
            DiagJp_ = "｜経路なし（" & nSec & "：呼出記録なし・ソース内参照も未検出→デッドコードの可能性）"
        End If
    ElseIf Left$(c, 9) = "conflict|" Then
        Dim body As String, q As Long, sfxJp As String
        body = Mid$(c, 10)
        q = InStrRev(body, "|")
        sfxJp = ""
        If q > 0 Then
            Select Case Mid$(body, q + 1)
                Case "tried"
                    sfxJp = "・転向/havoc 試行済"
                Case "nosite"
                    sfxJp = "・設値点なし"
                Case "nosteer"
                    sfxJp = "・転向情報なし（複合条件等）"
            End Select
            If Len(sfxJp) > 0 Then body = Left$(body, q - 1)
        End If
        DiagJp_ = "｜値競合: " & body & " を強制できず（定数伝播" & sfxJp & "）"
    ElseIf c = "dead" Then
        DiagJp_ = "｜経路構築不可"
    End If
End Function

Private Function TermsNote_(ByVal flow As OrderedDict) As String
    Dim s As String
    If flow.Exists("termsApplied") Then
        Dim ti As OrderedDict
        For Each ti In flow.Item("termsApplied")
            If Len(s) > 0 Then s = s & "、"
            s = s & CStr(ti.Item("name"))
            If CStr(ti.Item("source")) = "auto" Then
                s = s & "（自動検出）"
            Else
                s = s & "（コントロール登録）"
            End If
        Next ti
    End If
    If Len(s) = 0 Then
        TermsNote_ = "終了扱いセクション: なし　※名前に ABEND を含む SECTION は自動検出されます。他の命名はコントロールシート B24-B29 に登録してください"
    Else
        TermsNote_ = "終了扱いセクション: " & s
    End If
End Function

Private Function TermJp_(ByVal c As OrderedDict) As String
    Dim t As String
    t = CStr(c.Item("term"))
    If t = "goback" Then
        Dim tv As String
        If c.Exists("termVerb") Then tv = CStr(c.Item("termVerb"))
        If Len(tv) = 0 Then tv = "GOBACK"
        TermJp_ = tv & "（正常終了）"
    ElseIf Left$(t, 6) = "abend:" Then
        TermJp_ = "異常終了（" & Mid$(t, 7) & " 経由）"
    ElseIf Left$(t, 6) = "synth:" Then
        TermJp_ = "呼出先異常（CALL " & Mid$(t, 7) & "・合成）"
    Else
        TermJp_ = t
    End If
End Function

'======================================================================
' 分岐カバレッジ表 (matrix)
'======================================================================
Private Sub RenderMatrix_(ByVal flow As OrderedDict)
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(SHEET_MATRIX)
    ws.Cells.Clear
    ws.Columns(1).NumberFormat = "@"

    ws.Range("A1").Value = "分岐カバレッジ表（行 = 検証Point（分岐アーム） ／ 列 = テストケース ／ " & ChrW$(&H25CB) & " = 通過）"
    With ws.Range("A1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2").Value = "※ どのケースにも通過されない行（全空行）は赤 = 漏れ。恒真/恒偽の分岐（デッドコード）もここに現れます。"
    ws.Range("A2").Font.Color = RGB(120, 120, 120)
    ws.Range("A2").Font.Size = 9

    Dim cases As Collection
    Set cases = flow.Item("cases")
    Dim ncol As Long
    ncol = 3 + cases.Count    ' last data column (C = 備考, D.. = TCn)

    Dim hdr As Long
    hdr = 4
    ws.Cells(hdr, 1).Value = "検証Point（分岐アーム）"
    ws.Cells(hdr, 2).Value = "行"
    ws.Cells(hdr, 3).Value = "備考"
    Dim ci As Long, c As OrderedDict
    ci = 0
    For Each c In cases
        ci = ci + 1
        ws.Cells(hdr, 3 + ci).Value = CStr(c.Item("id"))
    Next c
    ' column header: dark blue band, white text; TC numbers centered
    With ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, ncol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 91, 143)
    End With
    ' each TC header is tinted by its 系: 正常=緑 / 異常=桃 (白字)
    ci = 0
    For Each c In cases
        ci = ci + 1
        ws.Cells(hdr, 3 + ci).Interior.Color = KindColor_(CStr(c.Item("kind")))
    Next c
    ws.Range(ws.Cells(hdr, 2), ws.Cells(hdr, ncol)).HorizontalAlignment = xlCenter

    ' body: SECTION bands (A=名称 / C=備考の漢字名) + 検証Point rows
    Dim row As Long, a As OrderedDict, hit As Boolean, anyHit As Boolean, v As Variant
    Dim secName As String, prevSec As String, blockStart As Long
    Dim blocks As Collection
    Set blocks = New Collection
    prevSec = ChrW$(1)    ' sentinel: no section yet
    row = hdr + 1
    For Each a In flow.Item("arms")
        secName = SectionOf_(flow, CLng(a.Item("Line")))
        If secName <> prevSec Then
            If prevSec <> ChrW$(1) Then blocks.Add Array(blockStart, row - 1)
            ws.Cells(row, 1).Value = secName
            ws.Cells(row, 3).Value = SectionNoteOf_(flow, secName)
            With ws.Range(ws.Cells(row, 1), ws.Cells(row, ncol))
                .Interior.Color = RGB(220, 230, 241)
                .Font.Bold = True
                .Font.Color = RGB(31, 78, 121)
            End With
            blockStart = row
            prevSec = secName
            row = row + 1
        End If
        ' zebra banding on alternate rows (overridden by the red NG fill)
        If ((row - hdr) Mod 2) = 0 Then
            ws.Range(ws.Cells(row, 1), ws.Cells(row, ncol)).Interior.Color = RGB(242, 244, 247)
        End If
        ws.Cells(row, 1).Value = CStr(a.Item("Disp"))
        On Error Resume Next
        ' HYPERLINK formula instead of Hyperlinks.Add: same click-to-jump,
        ' but a plain value write (much faster with many rows)
        ws.Cells(row, 2).Formula = "=HYPERLINK(""#'COBOLソース'!A" & (CLng(a.Item("Line")) + 3) & """," & CLng(a.Item("Line")) & ")"
        If Err.Number <> 0 Then
            ws.Cells(row, 2).Value = CLng(a.Item("Line"))
            Err.Clear
        End If
        On Error GoTo 0

        anyHit = False
        ci = 0
        For Each c In cases
            ci = ci + 1
            hit = False
            For Each v In c.Item("arms")
                If CStr(v) = CStr(a.Item("Token")) Then hit = True
            Next v
            If hit Then
                ws.Cells(row, 3 + ci).Value = ChrW$(&H25CB)
                ws.Cells(row, 3 + ci).HorizontalAlignment = xlCenter
                anyHit = True
            End If
        Next c
        If Not anyHit Then
            ws.Range(ws.Cells(row, 1), ws.Cells(row, ncol)).Interior.Color = RGB(255, 199, 206)
            ws.Cells(row, ncol + 1).Value = "未カバー" & DiagJp_(flow, CStr(a.Item("Token")))
            ws.Cells(row, ncol + 1).Font.Color = RGB(192, 0, 0)
        End If
        row = row + 1
    Next a
    If prevSec <> ChrW$(1) Then blocks.Add Array(blockStart, row - 1)

    If row > hdr + 1 Then
        With ws.Range(ws.Cells(hdr + 1, 2), ws.Cells(row - 1, 2))
            .Font.Color = RGB(5, 99, 193)
            .Font.Underline = True
            .HorizontalAlignment = xlCenter
        End With
        ' full grid (one range-level call - cheap regardless of row count)
        With ws.Range(ws.Cells(hdr, 1), ws.Cells(row - 1, ncol)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(184, 188, 196)
            .Weight = xlThin
        End With
    End If

    ' thick box per SECTION block (band + its rows)
    Dim bk As Variant
    For Each bk In blocks
        With ws.Range(ws.Cells(CLng(bk(0)), 1), ws.Cells(CLng(bk(1)), ncol))
            .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=RGB(89, 89, 89)
        End With
    Next bk

    ' footer: 系 / 終了形態 / 対象
    Dim ftrTop As Long
    row = row + 1
    ftrTop = row
    ws.Cells(row, 1).Value = "系"
    ws.Cells(row, 1).Font.Bold = True
    ci = 0
    For Each c In cases
        ci = ci + 1
        If CStr(c.Item("kind")) = "normal" Then
            ws.Cells(row, 3 + ci).Value = "正常"
        Else
            ws.Cells(row, 3 + ci).Value = "異常"
        End If
    Next c
    Dim sysRow As Long
    sysRow = row
    row = row + 1
    ws.Cells(row, 1).Value = "終了形態"
    ws.Cells(row, 1).Font.Bold = True
    ci = 0
    For Each c In cases
        ci = ci + 1
        ws.Cells(row, 3 + ci).Value = TermJp_(c)
        ws.Cells(row, 3 + ci).Font.Size = 9
    Next c
    row = row + 1
    ws.Cells(row, 1).Value = "対象"
    ws.Cells(row, 1).Font.Bold = True
    ci = 0
    For Each c In cases
        ci = ci + 1
        If CStr(c.Item("kind")) = "normal" Then
            ws.Cells(row, 3 + ci).Value = "テスト対象"
        Else
            ws.Cells(row, 3 + ci).Value = "機上確認"
            ws.Cells(row, 3 + ci).Font.Color = RGB(120, 120, 120)
        End If
    Next c

    ' footer block: light gray band + centered values + grid
    With ws.Range(ws.Cells(ftrTop, 1), ws.Cells(row, ncol))
        .Interior.Color = RGB(231, 233, 236)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(184, 188, 196)
        .Borders.Weight = xlThin
    End With
    ws.Range(ws.Cells(ftrTop, 4), ws.Cells(row, ncol)).HorizontalAlignment = xlCenter
    ' the 系 row reuses the TC header tint (緑/桃 + 白字) so 正常/異常 read at a glance
    ci = 0
    For Each c In cases
        ci = ci + 1
        With ws.Cells(sysRow, 3 + ci)
            .Interior.Color = KindColor_(CStr(c.Item("kind")))
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
    Next c

    ws.Columns("A").ColumnWidth = 56
    ws.Columns("B").ColumnWidth = 6
    ws.Columns("C").ColumnWidth = 24
    Dim k As Long
    For k = 1 To cases.Count + 1
        ws.Columns(3 + k).ColumnWidth = 12
    Next k

    ' whole-sheet font: Meiryo UI (Font.Name preserves bold/color/size)
    On Error Resume Next
    ws.UsedRange.Font.Name = "Meiryo UI"
    On Error GoTo 0
End Sub

' 系 tint: 正常 = 緑 / 異常 = 桃 (both with white text)
Private Function KindColor_(ByVal kind As String) As Long
    If kind = "normal" Then
        KindColor_ = RGB(84, 130, 53)
    Else
        KindColor_ = RGB(192, 80, 110)
    End If
End Function

' owning SECTION of a source line (from the flow result's section ranges)
Private Function SectionOf_(ByVal flow As OrderedDict, ByVal lineNo As Long) As String
    SectionOf_ = ""
    If Not flow.Exists("sections") Then Exit Function
    Dim s As OrderedDict
    For Each s In flow.Item("sections")
        If CLng(s.Item("line")) <= lineNo And lineNo <= CLng(s.Item("secEnd")) Then
            SectionOf_ = CStr(s.Item("name"))
        End If
    Next s
End Function

' the kanji description carried on a SECTION (the comment above its header)
Private Function SectionNoteOf_(ByVal flow As OrderedDict, ByVal name As String) As String
    SectionNoteOf_ = ""
    If Not flow.Exists("sections") Then Exit Function
    Dim s As OrderedDict
    For Each s In flow.Item("sections")
        If CStr(s.Item("name")) = name Then
            If s.Exists("note") Then SectionNoteOf_ = CStr(s.Item("note"))
            Exit Function
        End If
    Next s
End Function
