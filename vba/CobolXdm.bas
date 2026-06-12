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
