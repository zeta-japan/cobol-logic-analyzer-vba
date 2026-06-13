Attribute VB_Name = "CobolXdm"
' CobolXdm - XDM-style deliverable support (ver3.3)
'
' The client's test design lives in two sheets: a Japanese pattern table
' (one TC number per branch outcome) and 別紙1 (the source listing with
' TC numbers alongside the code). This module feeds both workflows:
'
'   ApplyTreeTc      - static TC column on ロジック階層(ソース順): every
'                      branch-arm row ([THEN]/[ELSE]/WHEN/[AT END]) gets
'                      the first test case that covers it. The tree render
'                      stamps each arm row's flow token into helper column
'                      F; this post-pass maps tokens to TC numbers into
'                      column D and clears the helper.
'   BuildBesshiDraft - 別紙1ドラフト sheet: A=行番号, B=ソース原文,
'                      E=ケースNo at each arm's mark line (the first
'                      statement inside the arm).
'
' One TC per 検証Point: the FIRST covering case wins (cases come ordered
' TC1, TC2, ... from the flow result).

Option Explicit

Private Const TREE_SHEET As String = "ロジック階層(ソース順)"
Private Const BESSHI_SHEET As String = "別紙1ドラフト"
Private Const COL_TC As Long = 4    ' tree sheet: TC number column
Private Const COL_TOK As Long = 6   ' tree sheet: helper token column
Private Const PATTERN_SHEET As String = "パターン表ドラフト"

' arm token -> id of the first covering case ("TC1", ...)
Public Function BuildTcMap(ByVal flowR As OrderedDict) As OrderedDict
    Dim map As OrderedDict
    Set map = New OrderedDict
    Set BuildTcMap = map
    If flowR Is Nothing Then Exit Function
    Dim c As OrderedDict, v As Variant
    For Each c In flowR.Item("cases")
        For Each v In c.Item("arms")
            If Not map.Exists(CStr(v)) Then map.Add CStr(v), CStr(c.Item("id"))
        Next v
    Next c
End Function

Public Sub ApplyTreeTc(ByVal flowR As OrderedDict)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(TREE_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    If lastRow < 11 Then Exit Sub

    Dim map As OrderedDict
    Set map = BuildTcMap(flowR)

    Dim r As Long, tok As String
    For r = 11 To lastRow
        tok = CStr(ws.Cells(r, COL_TOK).Value)
        If Len(tok) > 0 Then
            If map.Exists(tok) Then
                ws.Cells(r, COL_TC).Value = CStr(map.Item(tok))
            End If
        End If
    Next r
    ws.Columns(COL_TOK).ClearContents

    With ws.Columns(COL_TC)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Color = RGB(31, 78, 121)
        .ColumnWidth = 8
    End With
End Sub

Public Sub BuildBesshiDraft(ByVal flowR As OrderedDict, ByVal src As String)
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(BESSHI_SHEET)
    ws.Cells.Clear
    If flowR Is Nothing Then Exit Sub

    ' line number -> TC of the first covering arm at that mark line
    Dim map As OrderedDict, lineTc As OrderedDict, a As OrderedDict, k As String
    Set map = BuildTcMap(flowR)
    Set lineTc = New OrderedDict
    For Each a In flowR.Item("arms")
        If map.Exists(CStr(a.Item("Token"))) Then
            k = CStr(CLng(a.Item("MarkLine")))
            If Not lineTc.Exists(k) Then lineTc.Add k, CStr(map.Item(CStr(a.Item("Token"))))
        End If
    Next a

    ws.Cells(1, 1).Value = "行"
    ws.Cells(1, 2).Value = "修正後COBOL"
    ws.Cells(1, 5).Value = "ケース"
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, 5))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 91, 143)
    End With

    Dim lines() As String, n As Long, i As Long
    lines = Split(Replace(src, vbCrLf, vbLf), vbLf)
    n = UBound(lines) - LBound(lines) + 1
    ' a trailing newline yields one empty final element - drop it
    If n > 0 Then
        If Len(lines(UBound(lines))) = 0 Then n = n - 1
    End If
    If n < 1 Then Exit Sub

    ' bulk write: text format first so source lines are never formulas
    ws.Columns(2).NumberFormat = "@"
    Dim nums() As Variant, txts() As Variant
    ReDim nums(1 To n, 1 To 1)
    ReDim txts(1 To n, 1 To 1)
    For i = 1 To n
        nums(i, 1) = i
        txts(i, 1) = lines(LBound(lines) + i - 1)
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 1)).Value = nums
    ws.Range(ws.Cells(2, 2), ws.Cells(n + 1, 2)).Value = txts

    ' sparse TC marks
    Dim vKey As Variant
    For Each vKey In lineTc.Keys
        i = CLng(vKey)
        If i >= 1 And i <= n Then
            With ws.Cells(i + 1, 5)
                .Value = CStr(lineTc.Item(CStr(vKey)))
                .Font.Bold = True
                .Font.Color = RGB(31, 78, 121)
                .HorizontalAlignment = xlCenter
            End With
        End If
    Next vKey

    ws.Range(ws.Cells(2, 2), ws.Cells(n + 1, 2)).Font.Name = "ＭＳ ゴシック"
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, 1)).Font.Color = RGB(120, 120, 120)
    ws.Columns(1).ColumnWidth = 7
    ws.Columns(2).ColumnWidth = 90
    ws.Columns(5).ColumnWidth = 9
End Sub

' ============================ パターン表ドラフト ============================
' XDM「プログラムパターン表」の機械生成ドラフト: SECTION 順に全分岐を
' ①-1/①-2... で列挙し（直行文も処理行として出力）、条件をテンプレート日本語化、処理行は子アクション
' (PERFORM は直前コメントを採用) から要約、右端に ケースNo と分岐行番号。
' 自然な日本語への仕上げは人手で行う前提の「8割ドラフト」。


Public Sub BuildPatternDraft(ByVal flowR As OrderedDict, ByVal src As String)
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(PATTERN_SHEET)
    ws.Cells.Clear
    If flowR Is Nothing Then Exit Sub

    ' the same parse the flow ran (ids are line-derived, so tokens match)
    Dim norm As OrderedDict, nodes As Collection
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Set nodes = CobolParser.Get_CobolNodes(norm.Item("Lines"))

    ' comment text by line number (PERFORM の直前コメント参照用)
    Dim cmts As OrderedDict, ce As OrderedDict
    Set cmts = New OrderedDict
    For Each ce In norm.Item("Comments")
        cmts.Add CStr(ce.Item("Number")), CStr(ce.Item("Text"))
    Next ce

    Dim tcMap As OrderedDict
    Set tcMap = BuildTcMap(flowR)

    ws.Range("A1").Value = "パターン表ドラフト（機械生成 - 日本語は人手で仕上げてください）"
    With ws.Range("A1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2").Value = "※ 条件・処理はテンプレート変換（PERFORM は直前コメントを採用）。入れ子の分岐は親の枝の下にインデント表示（番号は出現順）。ケース列は分岐カバレッジ表と同じ TC 番号。"
    ws.Range("A2").Font.Color = RGB(120, 120, 120)
    ws.Range("A2").Font.Size = 9

    Dim row As Long
    row = 4
    ws.Cells(row, 1).Value = "No"
    ws.Cells(row, 2).Value = "番号"
    ws.Cells(row, 3).Value = "条件・処理"
    ws.Cells(row, 4).Value = "ケース"
    ws.Cells(row, 5).Value = "行"
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 5))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 91, 143)
    End With
    ws.Range(ws.Cells(row, 2), ws.Cells(row, 5)).HorizontalAlignment = xlCenter
    row = row + 1

    Dim secs As Collection, s As OrderedDict, secIdx As Long
    Set secs = flowR.Item("sections")
    For Each s In secs
        secIdx = secIdx + 1
        ' section band (+ そのすぐ上のコメントを補足として採用)
        ws.Cells(row, 1).Value = secIdx
        ws.Cells(row, 2).Value = CStr(s.Item("name"))
        Dim secNote As String
        secNote = NearbyComment_(cmts, CLng(s.Item("line")))
        If Len(secNote) > 0 Then ws.Cells(row, 3).Value = secNote
        With ws.Range(ws.Cells(row, 1), ws.Cells(row, 5))
            .Interior.Color = RGB(220, 230, 241)
            .Font.Bold = True
            .Font.Color = RGB(31, 78, 121)
        End With
        row = row + 1

        ' section body in SOURCE ORDER: straight-line statements become
        ' processing bullets, branches are numbered ①②... (nested branches
        ' render inline under their arm). A pure orchestration section -
        ' just a PERFORM sequence, no branches - now lists its steps too.
        Dim branchCtr As Long
        branchCtr = 0
        RenderSeq_ ws, row, nodes, CLng(s.Item("line")), CLng(s.Item("secEnd")), _
                   0, True, branchCtr, tcMap, cmts
    Next s

    If row > 5 Then
        With ws.Range(ws.Cells(4, 1), ws.Cells(row - 1, 5)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(184, 188, 196)
            .Weight = xlThin
        End With
    End If
    ws.Columns(1).ColumnWidth = 4
    ws.Columns(2).ColumnWidth = 16
    ws.Columns(3).ColumnWidth = 92
    ws.Columns(4).ColumnWidth = 8
    ws.Columns(5).ColumnWidth = 7
    ws.Columns(3).WrapText = False
End Sub

' render a node list in SOURCE ORDER. At the section top level topFilter
' keeps only nodes in [lo, hi]; nested (arm) calls pass topFilter=False
' since the children are already inside the section. depth drives the
' indentation; branchCtr (ByRef) numbers branches pre-order across the
' whole section so a nested branch follows its parent's number.
Private Sub RenderSeq_(ByVal ws As Worksheet, ByRef row As Long, ByVal list As Collection, _
                       ByVal lo As Long, ByVal hi As Long, ByVal depth As Long, _
                       ByVal topFilter As Boolean, ByRef branchCtr As Long, _
                       ByVal tcMap As OrderedDict, ByVal cmts As OrderedDict)
    If list Is Nothing Then Exit Sub
    Dim n As OrderedDict, t As String, ln As Long, inRange As Boolean, myNum As Long
    For Each n In list
        t = CStr(n.Item("type"))
        ln = CLng(n.Item("startLine"))
        inRange = True
        If topFilter Then
            If ln < lo Or ln > hi Then inRange = False
        End If
        If inRange Then
            If t = "if" Or t = "evaluate" Or t = "search" Then
                branchCtr = branchCtr + 1
                myNum = branchCtr
                RenderBranch_ ws, row, n, myNum, depth, branchCtr, tcMap, cmts
            Else
                EmitAction_ ws, row, n, depth, cmts
            End If
        End If
    Next n
End Sub

' one straight-line statement -> a 「・…」 processing row (skips the bare
' EXIT / CONTINUE that only terminate a paragraph)
Private Sub EmitAction_(ByVal ws As Worksheet, ByRef row As Long, ByVal n As OrderedDict, _
                        ByVal depth As Long, ByVal cmts As OrderedDict)
    Dim jp As String
    jp = ActionJp_(n, cmts)
    If Len(jp) = 0 Then Exit Sub
    ws.Cells(row, 3).Value = Indent_(depth) & jp
    row = row + 1
End Sub

' a branch -> numbered ①-1/①-2/... arm blocks; each arm recurses into its
' own body (nested branches inline, deeper indent)
Private Sub RenderBranch_(ByVal ws As Worksheet, ByRef row As Long, ByVal b As OrderedDict, _
                          ByVal myNum As Long, ByVal depth As Long, ByRef branchCtr As Long, _
                          ByVal tcMap As OrderedDict, ByVal cmts As OrderedDict)
    Dim t As String, num As String, ln As Long
    t = CStr(b.Item("type"))
    num = CircledNum_(myNum)
    ln = CLng(b.Item("startLine"))

    If t = "if" Then
        RenderArm_ ws, row, num & "-1", CondJp_(CStr(b.Item("condition"))) & " の場合", _
                   CStr(b.Item("id")) & ":then", b.Item("thenChildren"), ln, depth, branchCtr, tcMap, cmts
        Dim elLn As Long
        elLn = ln
        If Not IsNull(b.Item("elseLine")) Then elLn = CLng(b.Item("elseLine"))
        RenderArm_ ws, row, num & "-2", "上記以外の場合", _
                   CStr(b.Item("id")) & ":else", b.Item("elseChildren"), elLn, depth, branchCtr, tcMap, cmts
    ElseIf t = "evaluate" Then
        Dim cs As Collection, wi As Long, w As OrderedDict, hasOther As Boolean
        Set cs = b.Item("cases")
        For wi = 1 To cs.Count
            Set w = cs(wi)
            If CStr(w.Item("condition")) = "OTHER" Then
                hasOther = True
                RenderArm_ ws, row, num & "-" & wi, "上記以外の場合（WHEN OTHER）", _
                           CStr(w.Item("id")), w.Item("children"), CLng(w.Item("startLine")), depth, branchCtr, tcMap, cmts
            Else
                ' EVALUATE TRUE (DECIDE 変換形) は WHEN 条件そのものを表示
                Dim hd As String
                If UCase$(CStr(b.Item("expression"))) = "TRUE" Then
                    hd = CondJp_(CStr(w.Item("condition"))) & " の場合"
                Else
                    hd = CStr(b.Item("expression")) & " ＝ " & CStr(w.Item("condition")) & " の場合"
                End If
                RenderArm_ ws, row, num & "-" & wi, hd, _
                           CStr(w.Item("id")), w.Item("children"), CLng(w.Item("startLine")), depth, branchCtr, tcMap, cmts
            End If
        Next wi
        If Not hasOther Then
            RenderArm_ ws, row, num & "-" & (cs.Count + 1), "どの WHEN にも該当しない場合", _
                       CStr(b.Item("id")) & ":skip", Nothing, ln, depth, branchCtr, tcMap, cmts
        End If
    ElseIf t = "search" Then
        Dim sc As Collection, si As Long, sw As OrderedDict
        If Not IsNull(b.Item("atEndLine")) Then
            RenderArm_ ws, row, num & "-1", "検索該当なしの場合（AT END）", _
                       CStr(b.Item("id")) & ":atend", b.Item("atEndChildren"), CLng(b.Item("atEndLine")), depth, branchCtr, tcMap, cmts
        Else
            RenderArm_ ws, row, num & "-1", "検索該当なしの場合（AT END なし）", _
                       CStr(b.Item("id")) & ":skip", Nothing, ln, depth, branchCtr, tcMap, cmts
        End If
        Set sc = b.Item("cases")
        For si = 1 To sc.Count
            Set sw = sc(si)
            RenderArm_ ws, row, num & "-" & (si + 1), CondJp_(CStr(sw.Item("condition"))) & " の場合（SEARCH WHEN）", _
                       CStr(sw.Item("id")), sw.Item("children"), CLng(sw.Item("startLine")), depth, branchCtr, tcMap, cmts
        Next si
    End If
End Sub

' one arm = a condition row (番号 / 条件 / ケース / 行) followed by its body
' (rendered in source order one indent deeper); empty arms show 処理なし
Private Sub RenderArm_(ByVal ws As Worksheet, ByRef row As Long, ByVal num As String, _
                       ByVal condText As String, ByVal token As String, ByVal children As Collection, _
                       ByVal ln As Long, ByVal depth As Long, ByRef branchCtr As Long, _
                       ByVal tcMap As OrderedDict, ByVal cmts As OrderedDict)
    ws.Cells(row, 2).Value = num
    ws.Cells(row, 2).HorizontalAlignment = xlCenter
    ws.Cells(row, 3).Value = Indent_(depth) & condText
    ws.Cells(row, 3).Font.Bold = True
    If tcMap.Exists(token) Then
        ws.Cells(row, 4).Value = CStr(tcMap.Item(token))
        ws.Cells(row, 4).Font.Bold = True
        ws.Cells(row, 4).Font.Color = RGB(31, 78, 121)
    Else
        ws.Cells(row, 4).Value = "未カバー"
        ws.Cells(row, 4).Font.Color = RGB(192, 0, 0)
        ws.Cells(row, 4).Font.Size = 9
    End If
    ws.Cells(row, 4).HorizontalAlignment = xlCenter
    ws.Cells(row, 5).Value = ln
    ws.Cells(row, 5).Font.Color = RGB(120, 120, 120)
    row = row + 1

    Dim rowBefore As Long
    rowBefore = row
    RenderSeq_ ws, row, children, 0, 0, depth + 1, False, branchCtr, tcMap, cmts
    If row = rowBefore Then
        If children Is Nothing Then
            ws.Cells(row, 3).Value = Indent_(depth + 1) & "（処理なし）"
        Else
            ws.Cells(row, 3).Value = Indent_(depth + 1) & "（処理なし: CONTINUE）"
        End If
        ws.Cells(row, 3).Font.Color = RGB(120, 120, 120)
        row = row + 1
    End If
End Sub

' one statement label -> template Japanese (empty string = skip the row)
Private Function ActionJp_(ByVal n As OrderedDict, ByVal cmts As OrderedDict) As String
    ActionJp_ = ""
    If CStr(n.Item("type")) <> "action" Then
        ActionJp_ = "・条件分岐"
        Exit Function
    End If
    Dim lbl As String, ln As Long, p As Long, note As String
    lbl = CStr(n.Item("label"))
    ln = CLng(n.Item("startLine"))
    If lbl = "EXIT" Or lbl = "CONTINUE" Then Exit Function   ' terminator / no-op

    If Left$(lbl, 8) = "PERFORM " Then
        ' loop forms first (PERFORM UNTIL / VARYING / n TIMES)
        p = InStr(lbl, " UNTIL ")
        If p > 0 Then
            ActionJp_ = "・" & CondJp_(Trim$(Mid$(lbl, p + 7))) & " になるまで、以下を繰り返すこと（" & lbl & "）"
            Exit Function
        End If
        If InStr(lbl, " VARYING ") > 0 Then
            ActionJp_ = "・繰り返し処理を行うこと（" & lbl & "）"
            Exit Function
        End If
        If Len(lbl) > 6 Then
            If Right$(lbl, 6) = " TIMES" Then
                ActionJp_ = "・指定回数 繰り返すこと（" & lbl & "）"
                Exit Function
            End If
        End If
        ' simple PERFORM <para> (or THRU): adopt the comment above it
        note = NearbyComment_(cmts, ln)
        If Len(note) > 0 Then
            ActionJp_ = "・" & note & " を行うこと（" & lbl & "）"
        Else
            ActionJp_ = "・" & lbl & " を実行すること"
        End If
        Exit Function
    End If
    If Left$(lbl, 5) = "CALL " Then
        ActionJp_ = "・サブプログラム呼出を行うこと（" & lbl & "）"
        Exit Function
    End If
    If Left$(lbl, 6) = "GO TO " Then
        ActionJp_ = "・" & Trim$(Mid$(lbl, 7)) & " へ分岐すること"
        Exit Function
    End If
    If Left$(lbl, 5) = "MOVE " Then
        p = InStr(lbl, " TO ")
        If p > 0 Then
            ActionJp_ = "・" & Trim$(Mid$(lbl, p + 4)) & " に " & Trim$(Mid$(lbl, 6, p - 6)) & " を設定すること"
        Else
            ActionJp_ = "・" & lbl
        End If
        Exit Function
    End If
    If Left$(lbl, 4) = "ADD " Then
        p = InStr(lbl, " TO ")
        If p > 0 Then
            ActionJp_ = "・" & Trim$(Mid$(lbl, p + 4)) & " に " & Trim$(Mid$(lbl, 5, p - 5)) & " を加算すること"
            Exit Function
        End If
    ElseIf Left$(lbl, 9) = "SUBTRACT " Then
        p = InStr(lbl, " FROM ")
        If p > 0 Then
            ActionJp_ = "・" & Trim$(Mid$(lbl, p + 6)) & " から " & Trim$(Mid$(lbl, 10, p - 10)) & " を減算すること"
            Exit Function
        End If
    ElseIf Left$(lbl, 11) = "INITIALIZE " Then
        ActionJp_ = "・" & Trim$(Mid$(lbl, 12)) & " を初期化すること"
        Exit Function
    ElseIf Left$(lbl, 8) = "COMPUTE " Then
        ActionJp_ = "・" & Trim$(Mid$(lbl, 9)) & " を算出すること"
        Exit Function
    ElseIf Left$(lbl, 7) = "STRING " Then
        p = InStr(lbl, " INTO ")
        If p > 0 Then
            ActionJp_ = "・" & Trim$(Mid$(lbl, p + 6)) & " を編集（連結）すること"
        Else
            ActionJp_ = "・文字列編集を行うこと（" & lbl & "）"
        End If
        Exit Function
    ElseIf Left$(lbl, 5) = "READ " Then
        p = InStr(lbl, " INTO ")
        If p > 0 Then
            ActionJp_ = "・ファイルを読み込み " & Trim$(Mid$(lbl, p + 6)) & " に格納すること（" & lbl & "）"
        Else
            ActionJp_ = "・ファイルを読み込むこと（" & lbl & "）"
        End If
        Exit Function
    ElseIf Left$(lbl, 6) = "WRITE " Or Left$(lbl, 8) = "REWRITE " Then
        ActionJp_ = "・レコードを書き出すこと（" & lbl & "）"
        Exit Function
    End If
    If lbl = "GOBACK" Or Left$(lbl, 8) = "STOP RUN" Or Left$(lbl, 12) = "EXIT PROGRAM" Then
        ActionJp_ = "・処理を終了すること（" & lbl & "）"
        Exit Function
    End If
    ActionJp_ = "・" & Trunc_(lbl, 70)
End Function

' depth-based indentation (full-width spaces) for the 条件・処理 column
Private Function Indent_(ByVal depth As Long) As String
    Indent_ = ""
    Dim i As Long
    For i = 1 To depth
        Indent_ = Indent_ & ChrW$(&H3000)
    Next i
End Function

' comment on the line just above (up to 3 lines back, nearest wins)
Private Function NearbyComment_(ByVal cmts As OrderedDict, ByVal ln As Long) As String
    NearbyComment_ = ""
    Dim k As Long
    For k = ln - 1 To ln - 3 Step -1
        If k < 1 Then Exit For
        If cmts.Exists(CStr(k)) Then
            ' skip box-drawing rows (****...) - take a text comment only
            Dim ctext As String
            ctext = CStr(cmts.Item(CStr(k)))
            Dim bare As String
            bare = Replace(Replace(Replace(Replace(ctext, "*", ""), "-", ""), "=", ""), "/", "")
            bare = Replace(Replace(Replace(bare, "＊", ""), "―", ""), " ", "")
            If Len(bare) > 0 Then
                NearbyComment_ = Trim$(Replace(ctext, "*", ""))
                Exit Function
            End If
        Else
            Exit For   ' contiguous comments only
        End If
    Next k
End Function

' condition text -> template Japanese ("A ＝ 'x' または B ≠ 0")
Public Function CondJp_(ByVal cond As String) As String
    Dim s As String
    s = " " & cond & " "
    s = Replace(s, " NOT = ", " ≠ ")
    s = Replace(s, " = ", " ＝ ")
    s = Replace(s, " >= ", " ≧ ")
    s = Replace(s, " <= ", " ≦ ")
    s = Replace(s, " > ", " ＞ ")
    s = Replace(s, " < ", " ＜ ")
    s = Replace(s, " OR ", " または ")
    s = Replace(s, " AND ", " かつ ")
    CondJp_ = Trim$(s)
End Function

Private Function CircledNum_(ByVal n As Long) As String
    If n >= 1 And n <= 20 Then
        CircledNum_ = ChrW$(&H2460 + n - 1)
    Else
        CircledNum_ = "(" & n & ")"
    End If
End Function

Private Function Trunc_(ByVal s As String, ByVal n As Long) As String
    If Len(s) > n Then
        Trunc_ = Left$(s, n) & "..."
    Else
        Trunc_ = s
    End If
End Function

' test shim: route a synthetic action label through ActionJp_ (no comments)
Public Function ActionJpOf(ByVal label As String) As String
    Dim n As OrderedDict
    Set n = New OrderedDict
    n.Add "type", "action"
    n.Add "label", label
    n.Add "startLine", 1
    ActionJpOf = ActionJp_(n, New OrderedDict)
End Function
