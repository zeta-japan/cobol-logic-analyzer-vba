Attribute VB_Name = "CobolTcMark"
' CobolTcMark - ver2.1 feature: mark a test case (decision path) onto the
' ロジック階層 tree, in column D.
'
' Design (confirmed with the team):
'   - Test cases are CONSOLIDATED by their final Action (output result):
'     each path's group key = its LAST branch-governed action (an action that
'     is NOT executed by every path; the common GOBACK/EXIT tail is excluded).
'     ICASE1 / ICASE2 -> 96 paths collapse to 8 groups.
'   - One case is shown at a time. For the selected group, column D marks the
'     DECISION PATH: each branch row the representative path takes is annotated
'     with the arm (THEN / ELSE / WHEN / AT END), truncated at the final Action,
'     and the final Action row is emphasized (◆). Fall-through common rows are
'     left blank, so each case looks distinct.
'   - 前へ / 次へ buttons cycle through the groups.
'
' BuildTcGroups is pure logic (no sheet I/O) so it is unit-tested in Phase9.
' BuildTcMarking does the rendering and is called by Main after the tree is built.

Option Explicit

Private Const HIER_SHEET As String = "ロジック階層"
Private Const DATA_SHEET As String = "_TCData"
Private Const COL_TREE As Long = 1
Private Const COL_LINE As Long = 2
Private Const COL_KIND As Long = 3
Private Const COL_TC As Long = 4
Private Const COL_CTX As Long = 7   ' must match CobolExecTree.COL_CTX (branch context)
Private Const HEADER_TEXT As String = "階層ツリー"

' Tree background colors (must match CobolLogicViewer RenderHierarchy) so we can
' restore a row's original color after un-highlighting it.
Private Const C_BRANCH As Long = 16314338   ' RGB(226,239,248) branch rows
Private Const C_SECTION As Long = 13168895  ' RGB(255,240,200) SECTION headers
' Highlight colors for the selected case's covered path (column A:D).
Private Const HL_PATH As Long = 13561798    ' RGB(198,239,206) light green - covered branch
Private Const HL_FINAL As Long = 6740479    ' RGB(255,217,102) amber - final Action
Private Const HL_COVER As Long = 12909055   ' RGB(255,249,196) pale yellow - executed line (coverage)

' ===================================================================== '
'  Pure logic: build the consolidated test-case groups + decision paths  '
' ===================================================================== '
' tcs: a Collection of test-case OrderedDicts (root("testCases")). Each has
'   testCaseId (String), actionLines (Collection of Long),
'   actionLabels (Collection of String), branchIds (Collection of String).
' Returns a Collection of OrderedDict, one per group (sorted by finalLine):
'   finalLine (Long), finalLabel (String), repId (String),
'   memberCount (Long), decisionPath (String, e.g. "61:THEN 99:ELSE").
Public Function BuildTcGroups(ByVal tcs As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim total As Long
    total = tcs.Count
    If total = 0 Then Set BuildTcGroups = result: Exit Function

    Dim labelOf As OrderedDict, freq As OrderedDict
    Set labelOf = New OrderedDict
    Set freq = New OrderedDict

    Dim ti As Long, tc As OrderedDict, al As Collection, lbls As Collection, ai As Long, k As String
    For ti = 1 To total
        Set tc = tcs(ti)
        Set al = tc.Item("actionLines")
        Set lbls = tc.Item("actionLabels")
        For ai = 1 To al.Count
            labelOf.Add CStr(CLng(al(ai))), CStr(lbls(ai))
        Next ai
        ' count presence once per path (dedup within the path)
        Dim seen As OrderedDict
        Set seen = New OrderedDict
        For ai = 1 To al.Count
            k = CStr(CLng(al(ai)))
            If Not seen.Exists(k) Then
                seen.Add k, True
                If freq.Exists(k) Then
                    freq.Add k, CLng(freq.Item(k)) + 1
                Else
                    freq.Add k, CLng(1)
                End If
            End If
        Next ai
    Next ti

    ' group every path by its last branch-governed action line
    Dim groupKeys As OrderedDict
    Set groupKeys = New OrderedDict
    Dim fLine As Long, gk As String, col As Collection
    For ti = 1 To total
        Set tc = tcs(ti)
        Set al = tc.Item("actionLines")
        fLine = -1
        For ai = al.Count To 1 Step -1
            k = CStr(CLng(al(ai)))
            If CLng(freq.Item(k)) < total Then
                fLine = CLng(al(ai))
                Exit For
            End If
        Next ai
        gk = CStr(fLine)
        If groupKeys.Exists(gk) Then
            Set col = groupKeys.Item(gk)
        Else
            Set col = New Collection
            groupKeys.Add gk, col
        End If
        col.Add ti
    Next ti

    ' sort the group keys (final lines) ascending
    Dim keys As Collection
    Set keys = groupKeys.Keys
    Dim n As Long
    n = keys.Count
    If n = 0 Then Set BuildTcGroups = result: Exit Function
    Dim arr() As Long, ki As Long
    ReDim arr(1 To n)
    For ki = 1 To n
        arr(ki) = CLng(keys(ki))
    Next ki
    Dim a As Long, b As Long, tmp As Long
    For a = 2 To n
        tmp = arr(a)
        b = a - 1
        Do While b >= 1
            If arr(b) <= tmp Then Exit Do
            arr(b + 1) = arr(b)
            b = b - 1
        Loop
        arr(b + 1) = tmp
    Next a

    ' build one group record per final line
    For a = 1 To n
        fLine = arr(a)
        Set col = groupKeys.Item(CStr(fLine))
        ' representative = fewest branch decisions, tie-break lowest testCaseId
        Dim repIdx As Long, repBidN As Long, repId As String
        Dim ci As Long, cand As OrderedDict, bids As Collection, cBidN As Long, cId As String
        repIdx = 0: repBidN = 0: repId = ""
        For ci = 1 To col.Count
            Set cand = tcs(col(ci))
            Set bids = cand.Item("branchIds")
            cBidN = bids.Count
            cId = CStr(cand.Item("testCaseId"))
            If repIdx = 0 Or cBidN < repBidN Or (cBidN = repBidN And cId < repId) Then
                repIdx = col(ci): repBidN = cBidN: repId = cId
            End If
        Next ci

        ' decision path of the representative, truncated at the final line
        Set cand = tcs(repIdx)
        Set bids = cand.Item("branchIds")
        Dim dp As String, lnO As Long, armO As String
        dp = ""
        For ci = 1 To bids.Count
            ParseBranchId_ CStr(bids(ci)), lnO, armO
            If lnO > 0 And lnO <= fLine Then
                If Len(dp) > 0 Then dp = dp & " "
                dp = dp & CStr(lnO) & ":" & armO
            End If
        Next ci

        Dim flLabel As String
        If fLine >= 0 And labelOf.Exists(CStr(fLine)) Then
            flLabel = CStr(labelOf.Item(CStr(fLine)))
        Else
            flLabel = "(common)"
        End If

        ' full branchIds of the representative (for coverage highlighting) and
        ' its conditions (input setup, shown on the テストケース候補 sheet)
        Dim repBids As String, cv As Variant
        repBids = ""
        For ci = 1 To bids.Count
            If Len(repBids) > 0 Then repBids = repBids & ","
            repBids = repBids & CStr(bids(ci))
        Next ci
        Dim repConds As String, conds As Collection
        repConds = ""
        Set conds = cand.Item("conditions")
        For Each cv In conds
            If Len(repConds) > 0 Then repConds = repConds & " / "
            repConds = repConds & CStr(cv)
        Next cv

        Dim g As OrderedDict
        Set g = New OrderedDict
        g.Add "finalLine", fLine
        g.Add "finalLabel", flLabel
        g.Add "repId", repId
        g.Add "memberCount", col.Count
        g.Add "decisionPath", dp
        g.Add "repBranchIds", repBids
        g.Add "repConditions", repConds
        result.Add g
    Next a

    Set BuildTcGroups = result
End Function

' Parse a branchId ("if-99:then", "when-80", "search-121:atend") into a line
' number and an arm token (THEN / ELSE / WHEN / ATEND).
Private Sub ParseBranchId_(ByVal bid As String, ByRef lineOut As Long, ByRef armOut As String)
    lineOut = 0
    armOut = ""
    Dim p As Long, rest As String, q As Long, kindPart As String, numPart As String, armPart As String
    p = InStr(bid, "-")
    If p = 0 Then Exit Sub
    kindPart = Left$(bid, p - 1)
    rest = Mid$(bid, p + 1)
    q = InStr(rest, ":")
    If q > 0 Then
        numPart = Left$(rest, q - 1)
        armPart = Mid$(rest, q + 1)
    Else
        numPart = rest
        armPart = ""
    End If
    If IsNumeric(numPart) Then lineOut = CLng(numPart)
    Select Case True
        Case kindPart = "when"
            armOut = "WHEN"
        Case armPart = "then"
            armOut = "THEN"
        Case armPart = "else"
            armOut = "ELSE"
        Case armPart = "atend"
            armOut = "ATEND"
        Case Else
            armOut = UCase$(armPart)
    End Select
End Sub

' ===================================================================== '
'  Rendering: store groups, draw column-D header + buttons, mark group 1 '
' ===================================================================== '
' result: the OrderedDict returned by CobolParser.Analyze_Full (has testCases).
Public Sub BuildTcMarking(ByVal result As OrderedDict)
    On Error GoTo Done_
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(HIER_SHEET)

    ' Always clear any prior run's nav buttons + scratch data FIRST, so a file
    ' that yields no groups can never leave stale buttons driving MarkGroup_
    ' against a freshly rendered (different) tree.
    RemoveBtn_ ws, "btnTcPrev"
    RemoveBtn_ ws, "btnTcNext"
    ResetData_

    If result Is Nothing Then Exit Sub
    If Not result.Exists("testCases") Then Exit Sub

    Dim groups As Collection
    Set groups = BuildTcGroups(result.Item("testCases"))
    If groups.Count = 0 Then Exit Sub

    StoreGroups_ groups, HeaderRow_(ws)
    RenderGroupedTC_ groups      ' make テストケース候補 correspond to G1..Gn
    ws.Activate                 ' StoreGroups_ added/activated _TCData; bring the tree back to front
    HideData_                   ' now _TCData is not the active sheet -> very-hide succeeds
    DrawHeaderAndButtons_ ws, groups.Count
    MarkGroup_ 1
Done_:
End Sub

' Render the consolidated test cases (G1..Gn) onto the テストケース候補 sheet so
' its IDs match the ロジック階層 marking. Each row = one final-Action group with
' the representative's input conditions and the number of combinatorial paths it
' subsumes. (Overwrites the raw per-path list; the engine still enumerates all
' paths internally.)
Private Sub RenderGroupedTC_(ByVal groups As Collection)
    On Error GoTo GDone
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet("テストケース候補")
    If ws Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    ws.Cells.Clear
    ws.Columns("B:C").NumberFormat = "@"   ' COBOL text may contain '=' -> keep literal

    ws.Range("A1").Value = "テストケース候補（最終Actionで集約 ／ ロジック階層の G番号と対応）"
    With ws.Range("A1").Font
        .Bold = True: .Size = 13: .Color = RGB(38, 70, 83)
    End With

    Dim hdr As Long
    hdr = 3
    ws.Cells(hdr, 1).Value = "ケース"
    ws.Cells(hdr, 2).Value = "最終Action（出力結果）"
    ws.Cells(hdr, 3).Value = "入力条件（代表）"
    ws.Cells(hdr, 4).Value = "該当パス数"
    ws.Cells(hdr, 5).Value = "最終Action行"
    With ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 5))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 232)
    End With

    Dim i As Long, g As OrderedDict, row As Long
    row = hdr + 1
    For i = 1 To groups.Count
        Set g = groups(i)
        ws.Cells(row, 1).Value = "G" & i
        ws.Cells(row, 1).Font.Bold = True
        ws.Cells(row, 2).Value = CStr(g.Item("finalLabel"))
        ws.Cells(row, 3).Value = CStr(g.Item("repConditions"))
        ws.Cells(row, 3).WrapText = True
        ws.Cells(row, 4).Value = CLng(g.Item("memberCount"))
        ws.Cells(row, 5).Value = CLng(g.Item("finalLine"))
        row = row + 1
    Next i

    ws.Cells(row + 1, 1).Value = "※ ロジック階層シートの「前へ／次へ」で各 G の決定パス・実行行（カバレッジ）を可視化できます。"
    ws.Cells(row + 1, 1).Font.Color = RGB(120, 120, 120)

    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 64
    ws.Columns("D").ColumnWidth = 12
    ws.Columns("E").ColumnWidth = 12
    Application.ScreenUpdating = True
GDone:
    Application.ScreenUpdating = True
End Sub

Public Sub TC_ShowNext()
    StepGroup_ 1
End Sub

Public Sub TC_ShowPrev()
    StepGroup_ -1
End Sub

Private Sub StepGroup_(ByVal delta As Long)
    On Error GoTo Done_
    Dim wd As Worksheet
    Set wd = DataSheet_(False)
    If wd Is Nothing Then Exit Sub
    Dim cnt As Long
    cnt = CLng(wd.Cells(2, 1).Value)
    If cnt = 0 Then Exit Sub
    Dim idx As Long
    idx = CLng(wd.Cells(1, 1).Value) + delta
    If idx < 1 Then idx = cnt
    If idx > cnt Then idx = 1
    MarkGroup_ idx
Done_:
End Sub

Private Function HeaderRow_(ByVal ws As Worksheet) As Long
    Dim r As Long
    For r = 1 To 40
        If InStr(1, CStr(ws.Cells(r, COL_TREE).Value), HEADER_TEXT) > 0 Then
            HeaderRow_ = r
            Exit Function
        End If
    Next r
    HeaderRow_ = 10 ' fallback to the known layout
End Function

Private Function DataSheet_(ByVal createIfMissing As Boolean) As Worksheet
    Dim wd As Worksheet
    On Error Resume Next
    Set wd = ThisWorkbook.Sheets(DATA_SHEET)
    On Error GoTo 0
    If wd Is Nothing And createIfMissing Then
        Set wd = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wd.Name = DATA_SHEET
    End If
    Set DataSheet_ = wd
End Function

' Very-hide the scratch sheet. Call only when _TCData is NOT the active sheet
' (Excel raises 1004 when you try to hide the active sheet).
Private Sub HideData_()
    Dim wd As Worksheet
    On Error Resume Next
    Set wd = ThisWorkbook.Sheets(DATA_SHEET)
    If Not wd Is Nothing Then wd.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' Reset the scratch sheet's group count to 0 so stale data can never drive
' MarkGroup_ against a newly rendered tree.
Private Sub ResetData_()
    Dim wd As Worksheet
    On Error Resume Next
    Set wd = ThisWorkbook.Sheets(DATA_SHEET)
    On Error GoTo 0
    If Not wd Is Nothing Then wd.Cells(2, 1).Value = 0
End Sub

Private Sub StoreGroups_(ByVal groups As Collection, ByVal hdrRow As Long)
    Dim wd As Worksheet
    Set wd = DataSheet_(True)
    If wd Is Nothing Then Exit Sub
    wd.Cells.Clear
    wd.Columns("A:G").NumberFormat = "@"
    wd.Cells(1, 1).Value = 1            ' current index
    wd.Cells(2, 1).Value = groups.Count ' group count
    wd.Cells(1, 2).Value = hdrRow       ' header row of the tree
    Dim i As Long, g As OrderedDict
    For i = 1 To groups.Count
        Set g = groups(i)
        wd.Cells(2 + i, 1).Value = g.Item("finalLine")
        wd.Cells(2 + i, 2).Value = g.Item("finalLabel")
        wd.Cells(2 + i, 3).Value = g.Item("repId")
        wd.Cells(2 + i, 4).Value = g.Item("memberCount")
        wd.Cells(2 + i, 5).Value = g.Item("decisionPath")
        wd.Cells(2 + i, 6).Value = g.Item("repBranchIds")
        wd.Cells(2 + i, 7).Value = g.Item("repConditions")
    Next i
End Sub

Private Sub DrawHeaderAndButtons_(ByVal ws As Worksheet, ByVal groupCount As Long)
    Dim hdrRow As Long
    hdrRow = HeaderRow_(ws)

    ' column-D header, matching the A:C tree header band
    ws.Cells(hdrRow, COL_TC).Value = "TestCase"
    ws.Cells(hdrRow, COL_TC).Font.Bold = True
    ws.Cells(hdrRow, COL_TC).Interior.Color = RGB(217, 225, 232)
    ws.Columns(COL_TC).ColumnWidth = 16

    ' control panel above the tree (rows 3-6, columns D onward are free)
    ws.Range("D3").Value = ChrW$(&H25A0) & " テストケース標記（決定パス）"
    ws.Range("D3").Font.Bold = True
    ws.Range("D3").Font.Color = RGB(38, 70, 83)

    ' (re)create the navigation buttons; delete old ones first so a re-run
    ' does not stack duplicates (Cells.Clear does not remove shapes).
    RemoveBtn_ ws, "btnTcPrev"
    RemoveBtn_ ws, "btnTcNext"
    AddBtn_ ws, "btnTcPrev", "<< 前のケース", ws.Range("D6"), "TC_ShowPrev"
    AddBtn_ ws, "btnTcNext", "次のケース >>", ws.Range("F6"), "TC_ShowNext"
End Sub

Private Sub RemoveBtn_(ByVal ws As Worksheet, ByVal nm As String)
    On Error Resume Next
    ws.Shapes(nm).Delete
    On Error GoTo 0
End Sub

Private Sub AddBtn_(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, _
                    ByVal anchor As Range, ByVal macro As String)
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, anchor.Left, anchor.Top, 110, 26) ' 5 = RoundedRectangle
    btn.Name = nm
    btn.Fill.ForeColor.RGB = RGB(42, 157, 143)
    btn.Line.Visible = msoFalse
    With btn.TextFrame2.TextRange
        .Text = caption
        .ParagraphFormat.Alignment = msoAlignCenter
        .Font.Size = 10
        .Font.Bold = msoTrue
        .Font.Name = "Meiryo UI"
        .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.OnAction = macro
End Sub

' True if the 種別 (column C) text denotes a branch row that RenderHierarchy
' painted with C_BRANCH (so we restore that color when un-highlighting).
Private Function IsBranchKind_(ByVal kindText As String) As Boolean
    Select Case kindText
        Case "IF", "EVALUATE", "SEARCH", "WHEN", "BRANCH"
            IsBranchKind_ = True
        Case Else
            IsBranchKind_ = False
    End Select
End Function

' A tree row is executed by the representative path iff every branchId token of
' its branch context (CobolExecTree column COL_CTX) is taken by that path. An
' empty context = always-executed (the spine), so it is covered.
Private Function IsCovered_(ByVal ctx As String, ByVal repSet As OrderedDict) As Boolean
    If Len(ctx) = 0 Then
        IsCovered_ = True
        Exit Function
    End If
    Dim parts() As String, i As Long
    parts = Split(ctx, "|")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            If Not repSet.Exists(parts(i)) Then
                IsCovered_ = False
                Exit Function
            End If
        End If
    Next i
    IsCovered_ = True
End Function

Private Sub MarkGroup_(ByVal idx As Long)
    On Error GoTo Done_
    Dim wd As Worksheet
    Set wd = DataSheet_(False)
    If wd Is Nothing Then Exit Sub
    Dim cnt As Long
    cnt = CLng(wd.Cells(2, 1).Value)
    If cnt = 0 Then Exit Sub
    If idx < 1 Then idx = cnt
    If idx > cnt Then idx = 1
    wd.Cells(1, 1).Value = idx

    Dim hdrRow As Long
    hdrRow = CLng(wd.Cells(1, 2).Value)
    If hdrRow < 1 Then hdrRow = 10
    Dim treeStart As Long
    treeStart = hdrRow + 1

    Dim r As Long
    r = 2 + idx
    Dim finalLine As Long, finalLabel As String, memberCount As Long, dpEnc As String
    finalLine = CLng(wd.Cells(r, 1).Value)
    finalLabel = CStr(wd.Cells(r, 2).Value)
    memberCount = CLng(wd.Cells(r, 4).Value)
    dpEnc = CStr(wd.Cells(r, 5).Value)

    ' decode decision path into a line -> arm map
    Dim armOf As OrderedDict
    Set armOf = New OrderedDict
    If Len(dpEnc) > 0 Then
        Dim parts() As String, i As Long, kv() As String
        parts = Split(dpEnc, " ")
        For i = LBound(parts) To UBound(parts)
            If Len(parts(i)) > 0 Then
                kv = Split(parts(i), ":")
                If UBound(kv) >= 1 Then armOf.Add kv(0), kv(1)
            End If
        Next i
    End If

    ' representative's full branchIds -> set, for coverage highlighting: a tree
    ' row is "executed" by this case iff every token of its branch context is in
    ' this set.
    Dim repSet As OrderedDict
    Set repSet = New OrderedDict
    Dim rbi As String
    rbi = CStr(wd.Cells(r, 6).Value)
    If Len(rbi) > 0 Then
        Dim bparts() As String, bi As Long
        bparts = Split(rbi, ",")
        For bi = LBound(bparts) To UBound(bparts)
            If Len(bparts(bi)) > 0 And Not repSet.Exists(bparts(bi)) Then repSet.Add bparts(bi), True
        Next bi
    End If

    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(HIER_SHEET)
    Application.ScreenUpdating = False

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TREE).End(xlUp).Row
    If lastRow < treeStart Then lastRow = treeStart

    Dim row As Long, lnVal As Variant, key As String, arm As String
    Dim isFinal As Boolean, cellText As String, colA As String, colC As String
    Dim isSectionHdr As Boolean, covered As Boolean
    For row = treeStart To lastRow
        ' (1) restore this row's A:C background to its original tree color and
        '     clear the previous case's column-D mark.
        colA = CStr(ws.Cells(row, COL_TREE).Value)
        colC = CStr(ws.Cells(row, COL_KIND).Value)
        isSectionHdr = (InStr(colA, ChrW$(&H25A0)) > 0 And InStr(colA, "SECTION") > 0)
        If isSectionHdr Then
            ws.Range(ws.Cells(row, COL_TREE), ws.Cells(row, COL_KIND)).Interior.Color = C_SECTION
        ElseIf IsBranchKind_(colC) Then
            ws.Range(ws.Cells(row, COL_TREE), ws.Cells(row, COL_KIND)).Interior.Color = C_BRANCH
        Else
            ws.Range(ws.Cells(row, COL_TREE), ws.Cells(row, COL_KIND)).Interior.ColorIndex = xlNone
        End If
        ws.Cells(row, COL_TC).Clear

        ' (2) is this row executed by the selected case's representative path?
        covered = False
        If Not isSectionHdr Then covered = IsCovered_(CStr(ws.Cells(row, COL_CTX).Value), repSet)

        ' (3) decision arm / final action (rows that carry a source line)
        lnVal = ws.Cells(row, COL_LINE).Value
        isFinal = False
        arm = ""
        If IsNumeric(lnVal) And Len(CStr(lnVal)) > 0 Then
            key = CStr(CLng(lnVal))
            isFinal = (CLng(lnVal) = finalLine)
            If armOf.Exists(key) Then
                arm = CStr(armOf.Item(key))
                If arm = "ATEND" Then arm = "AT END"
            End If
        End If

        ' (4) paint: final (amber) > decision arm (green) > executed (pale yellow)
        If isFinal Then
            cellText = arm
            If Len(cellText) > 0 Then cellText = cellText & " "
            cellText = cellText & ChrW$(&H25C6) & " 最終Action"
            ws.Cells(row, COL_TC).Value = cellText
            ws.Cells(row, COL_TC).Font.Bold = True
            ws.Cells(row, COL_TC).Font.Color = RGB(192, 0, 0)
            ws.Range(ws.Cells(row, COL_TREE), ws.Cells(row, COL_TC)).Interior.Color = HL_FINAL
        ElseIf arm <> "" Then
            ws.Cells(row, COL_TC).Value = arm
            ws.Cells(row, COL_TC).Font.Bold = True
            ws.Cells(row, COL_TC).Font.Color = RGB(0, 112, 192)
            ws.Range(ws.Cells(row, COL_TREE), ws.Cells(row, COL_TC)).Interior.Color = HL_PATH
        ElseIf covered Then
            ws.Range(ws.Cells(row, COL_TREE), ws.Cells(row, COL_TC)).Interior.Color = HL_COVER
        End If
    Next row

    ' status line
    ws.Range("D4").Value = "ケース G" & idx & "/" & cnt & " : " & finalLabel & _
        "  ｜ " & memberCount & "パス相当  ｜ 凡例 " & ChrW$(&H25C6) & "=最終Action / 緑=分岐選択 / 薄黄=実行行（カバレッジ）"
    ws.Range("D4").Font.Color = RGB(80, 80, 80)

    Application.ScreenUpdating = True
Done_:
    Application.ScreenUpdating = True
End Sub
