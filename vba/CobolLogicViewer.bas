Attribute VB_Name = "CobolLogicViewer"
' =========================================================================
' CobolLogicViewer - render COBOL analysis JSON into 5 Excel sheets.
'
' Reads a JSON file (produced by the in-workbook engine) and builds:
'   0. COBOL source   : original lines with workbook row numbers
'   1. Logic tree     : IF / EVALUATE / SEARCH nested structure with section headers
'   2. Test cases     : test case table generated from enumerated execution paths
'   3. Branch coverage: per-branch coverage summary
'   4. Call graph     : PERFORM / CALL / section / paragraph call relationships
'
' Public entry points:
'   SetupControlSheet      : create the Control sheet and button (one-time)
'   PickJsonAndBuild       : show file picker -> BuildCobolReport
'   BuildCobolReport(path) : parse the JSON at path and build all 5 sheets
'
' Dependencies: JsonParser module (ParseJson / ReadAllText / EnsureSheet),
'               Reference to Microsoft Scripting Runtime.
' =========================================================================
Option Explicit

' 色定数
Private Const C_HEADER  As Long = 14474460   ' RGB(220,220,220)
Private Const C_OK      As Long = 13561798   ' RGB(198,239,206)
Private Const C_WARN    As Long = 10284031   ' RGB(255,235,156)
Private Const C_ERR     As Long = 13551615   ' RGB(255,199,206)
Private Const C_BRANCH  As Long = 16314338   ' RGB(226,239,248)
Private Const C_SECTION As Long = 13168895   ' RGB(255,240,200) SECTION ヘッダ

'==========================================================================
' 公開エントリポイント
'==========================================================================

Private Sub LegacySetupControlSheet_()   ' retired: Main.SetupControlSheet is the real one
    Dim ws As Worksheet
    Set ws = EnsureSheet("コントロール")
    ws.Cells.Clear

    ws.Range("A1").Value = "COBOL ロジック解析ツール"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3").Value = "■ 使い方"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").Value = "   1. 下のボタンで COBOL ソース (.cbl) を選択"
    ws.Range("A5").Value = "   2. 解析結果が 5 つのシートに自動生成されます"
    ws.Range("A6").Value = "      (COBOLソース / ロジック階層 / テストケース / 分岐カバレッジ / 呼出関係)"

    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0

    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("A8").Left, ws.Range("A8").Top, 280, 44)
    btn.OnAction = "PickJsonAndBuild"
    btn.Caption = "JSON を読み込んでレポート生成"
    btn.Font.Size = 11
    btn.Font.Bold = True

    ws.Columns("A").ColumnWidth = 80

    ws.Activate
    ws.Range("A1").Select
    MsgBox "コントロール シートを作成しました。" & vbLf & _
           "今後はこのシートのボタンから利用できます。"
End Sub

Public Sub PickJsonAndBuild()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "解析結果 JSON を選択"
    fd.Filters.Clear
    fd.Filters.Add "ロジック解析 JSON", "*.logic.json"
    fd.Filters.Add "All files", "*.*"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        BuildCobolReport CStr(fd.SelectedItems(1))
    End If
End Sub

Public Sub BuildCobolReport(ByVal logicJsonPath As String, Optional ByVal showDone As Boolean = True)
    On Error GoTo Fail
    Dim root As Object
    Set root = ParseJson(ReadAllText(logicJsonPath))
    Application.ScreenUpdating = False

    ' COBOLソース を先に作成 (ロジック階層からハイパーリンクするため)
    RenderSource    EnsureSheet("COBOLソース"),       root
    RenderHierarchy EnsureSheet("ロジック階層(ソース順)"), root
    CobolExecTree.RenderExecHierarchy EnsureSheet("ロジック階層(実行順展開)"), root
    RenderCoverage  EnsureSheet("分岐カバレッジ"),   root
    RenderCallGraph EnsureSheet("呼出関係"),         root

    ThisWorkbook.Sheets("ロジック階層(実行順展開)").Activate
    Application.ScreenUpdating = True

    If showDone Then MsgBox "解析レポート生成完了" & vbLf & logicJsonPath
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    MsgBox "レポート生成中にエラーが発生しました: " & Err.Description, vbExclamation
End Sub

'==========================================================================
' シート: COBOLソース
'==========================================================================

Private Sub RenderSource(ws As Worksheet, root As Object)
    ws.Cells.Clear
    Dim src As Object
    Set src = root("source")
    If src Is Nothing Then Exit Sub
    If src.Count = 0 Then Exit Sub

    ws.Range("A1").Value = "COBOLソース"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "(行番号 = ファイル上の物理行)"

    ws.Cells(3, 1).Value = "行番号"
    ws.Cells(3, 2).Value = "ソース"
    ws.Range("A3:B3").Font.Bold = True
    ws.Range("A3:B3").Interior.Color = C_HEADER

    Dim n As Long
    n = src.Count
    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        arr(i, 1) = i
        arr(i, 2) = CStr(src(i))
    Next i
    ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 2)).Value = arr

    ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 2)).Font.Name = "MS Gothic"
    ws.Range(ws.Cells(4, 2), ws.Cells(3 + n, 2)).Font.Size = 10
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 100

    ws.Activate
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
End Sub

'==========================================================================
' シート: ロジック階層
'==========================================================================

Private Sub RenderHierarchy(ws As Worksheet, root As Object)
    ws.Cells.Clear
    Dim s As Object
    Set s = root("summary")

    ws.Range("A1").Value = "COBOLロジック階層（ソース順）"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "プログラム名"
    ws.Range("B3").Value = s("programName")
    ws.Range("A4").Value = "行数"
    ws.Range("B4").Value = s("lines")
    ws.Range("A5").Value = "分岐数"
    ws.Range("B5").Value = s("branchCount")
    ws.Range("A6").Value = "パス数"
    ws.Range("B6").Value = s("pathCount")
    ws.Range("A7").Value = "プレフィックス形式"
    ws.Range("B7").Value = s("prefixStyle") & " (検出: " & s("prefixDetected") & ")"
    ws.Range("A8").Value = "警告件数"
    ws.Range("B8").Value = root("warnings").Count

    Dim row As Long
    row = 10
    ws.Cells(row, 1).Value = "階層ツリー"
    ws.Cells(row, 2).Value = "開始行"
    ws.Cells(row, 3).Value = "種別"
    ws.Cells(row, 4).Value = "TC"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Interior.Color = C_HEADER
    row = row + 1

    Dim treeStart As Long
    treeStart = row

    Dim nodes As Object
    Set nodes = root("rootNodes")
    Dim sections As Object
    Set sections = root("programStructure")("sections")

    Dim currentSection As String
    currentSection = "###NONE###"

    Dim i As Long
    For i = 1 To nodes.Count
        Dim node As Object
        Set node = nodes(i)
        Dim nodeSection As String
        nodeSection = FindSectionForLine(sections, CLng(node("startLine")))

        If nodeSection <> currentSection Then
            If nodeSection <> "" Then
                ws.Cells(row, 1).Value = ChrW$(&H25A0) & " " & nodeSection & " SECTION"
                ws.Cells(row, 1).Font.Bold = True
                ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Interior.Color = C_SECTION
                row = row + 1
            End If
            currentSection = nodeSection
        End If

        Dim isLastInSection As Boolean
        isLastInSection = True
        If i < nodes.Count Then
            If FindSectionForLine(sections, CLng(nodes(i + 1)("startLine"))) = currentSection Then
                isLastInSection = False
            End If
        End If

        RenderNode ws, node, "", isLastInSection, row
    Next i

    Dim treeEnd As Long
    treeEnd = row - 1
    If treeEnd >= treeStart Then
        ws.Range(ws.Cells(treeStart, 1), ws.Cells(treeEnd, 1)).Font.Name = "MS Gothic"
        With ws.Range(ws.Cells(treeStart, 2), ws.Cells(treeEnd, 2)).Font
            .Color = RGB(5, 99, 193)
            .Underline = True
        End With
    End If

    RenderWarnings ws, root, row
    ws.Columns("A:C").AutoFit
    If ws.Columns("A").ColumnWidth > 90 Then ws.Columns("A").ColumnWidth = 90
End Sub

Private Sub RenderNode(ws As Worksheet, node As Object, ByVal prefix As String, _
                       ByVal isLast As Boolean, ByRef row As Long)
    Dim connector As String, childPrefix As String
    If isLast Then
        connector = prefix & ChrW$(&H2514) & ChrW$(&H2500) & " "
        childPrefix = prefix & "    "
    Else
        connector = prefix & ChrW$(&H251C) & ChrW$(&H2500) & " "
        childPrefix = prefix & ChrW$(&H2502) & "   "
    End If

    Dim t As String, label As String, kind As String
    t = node("type")
    Select Case t
        Case "if"
            label = "IF " & node("condition"): kind = "IF"
        Case "evaluate"
            label = "EVALUATE " & node("expression"): kind = "EVALUATE"
        Case "search"
            label = "SEARCH " & IIf(node("isAll"), "ALL ", "") & node("tableExpr"): kind = "SEARCH"
        Case "when"
            label = "WHEN " & node("condition"): kind = "WHEN"
        Case "action"
            label = node("label"): kind = "ACTION"
        Case Else
            label = t: kind = t
    End Select

    ws.Cells(row, 1).Value = connector & label

    ' 開始行: COBOLソース シートへハイパーリンク
    Dim sl As Long
    sl = CLng(node("startLine"))
    On Error Resume Next
    ws.Cells(row, 2).Formula = "=HYPERLINK(""#'COBOLソース'!A" & (sl + 3) & """," & sl & ")"
    If Err.Number <> 0 Then
        ws.Cells(row, 2).Value = sl
        Err.Clear
    End If
    On Error GoTo 0

    ws.Cells(row, 3).Value = kind
    ' arm rows carry their flow-arm token in a helper column; a post-pass
    ' (CobolXdm.ApplyTreeTc) maps tokens to TC numbers and clears it
    If t = "when" Then ws.Cells(row, 6).Value = CStr(node("id"))
    If t <> "action" Then
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Interior.Color = C_BRANCH
    End If
    row = row + 1

    Select Case t
        Case "if"
            Dim elseCount As Long
            elseCount = node("elseChildren").Count
            RenderBranchGroup ws, "[THEN]", node("thenChildren"), childPrefix, (elseCount = 0), row, CStr(node("id")) & ":then"
            If elseCount > 0 Then
                RenderBranchGroup ws, "[ELSE]", node("elseChildren"), childPrefix, True, row, CStr(node("id")) & ":else"
            End If
        Case "evaluate"
            Dim cs As Object, ci As Long
            Set cs = node("cases")
            For ci = 1 To cs.Count
                RenderNode ws, cs(ci), childPrefix, (ci = cs.Count), row
            Next ci
        Case "search"
            Dim atEnd As Object, sc As Object, si As Long
            Set atEnd = node("atEndChildren")
            Set sc = node("cases")
            If atEnd.Count > 0 Then
                RenderBranchGroup ws, "[AT END]", atEnd, childPrefix, (sc.Count = 0), row, CStr(node("id")) & ":atend"
            End If
            For si = 1 To sc.Count
                RenderNode ws, sc(si), childPrefix, (si = sc.Count), row
            Next si
        Case "when"
            Dim wc As Object, wi As Long
            Set wc = node("children")
            For wi = 1 To wc.Count
                RenderNode ws, wc(wi), childPrefix, (wi = wc.Count), row
            Next wi
    End Select
End Sub

Private Sub RenderBranchGroup(ws As Worksheet, ByVal groupLabel As String, children As Object, _
                              ByVal prefix As String, ByVal isLast As Boolean, ByRef row As Long, _
                              ByVal armToken As String)
    Dim connector As String, childPrefix As String
    If isLast Then
        connector = prefix & ChrW$(&H2514) & ChrW$(&H2500) & " "
        childPrefix = prefix & "    "
    Else
        connector = prefix & ChrW$(&H251C) & ChrW$(&H2500) & " "
        childPrefix = prefix & ChrW$(&H2502) & "   "
    End If

    ws.Cells(row, 1).Value = connector & groupLabel
    ws.Cells(row, 3).Value = "BRANCH"
    If Len(armToken) > 0 Then ws.Cells(row, 6).Value = armToken
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Interior.Color = C_BRANCH
    row = row + 1

    Dim i As Long
    For i = 1 To children.Count
        RenderNode ws, children(i), childPrefix, (i = children.Count), row
    Next i
End Sub

Private Function FindSectionForLine(sections As Object, ByVal lineNum As Long) As String
    FindSectionForLine = ""
    If sections Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To sections.Count
        If CLng(sections(i)("line")) <= lineNum Then
            FindSectionForLine = CStr(sections(i)("name"))
        Else
            Exit For
        End If
    Next i
End Function

'==========================================================================
' シート: テストケース候補
'==========================================================================

Private Sub RenderTestCases(ws As Worksheet, root As Object)
    ws.Cells.Clear
    ws.Range("A1").Value = "テストケース候補"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim hdr As Long
    hdr = 3
    ws.Cells(hdr, 1).Value = "TC番号"
    ws.Cells(hdr, 2).Value = "シナリオ名"
    ws.Cells(hdr, 3).Value = "入力条件"
    ws.Cells(hdr, 4).Value = "期待アクション"
    ws.Cells(hdr, 5).Value = "優先度"
    ws.Cells(hdr, 6).Value = "対象行"
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 6)).Font.Bold = True
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 6)).Interior.Color = C_HEADER

    Dim tcs As Object
    Set tcs = root("testCases")
    Dim row As Long, i As Long
    row = hdr + 1
    For i = 1 To tcs.Count
        Dim tc As Object
        Set tc = tcs(i)
        ws.Cells(row, 1).Value = tc("testCaseId")
        ws.Cells(row, 2).Value = tc("scenarioName")
        ws.Cells(row, 3).Value = JoinList(tc("conditions"), " / ")
        ws.Cells(row, 4).Value = JoinList(tc("actionLabels"), vbLf)
        ws.Cells(row, 5).Value = tc("priority")
        ws.Cells(row, 6).Value = tc("sourceLineText")
        ws.Cells(row, 4).WrapText = True
        ws.Cells(row, 3).WrapText = True
        If CStr(tc("priority")) = "高" Then
            ws.Range(ws.Cells(row, 1), ws.Cells(row, 6)).Font.Bold = True
        End If
        row = row + 1
    Next i

    ws.Cells(row + 1, 1).Value = "総テストケース数: " & tcs.Count
    ws.Cells(row + 1, 1).Font.Bold = True
    row = row + 1

    RenderWarnings ws, root, row
    ws.Columns("A:F").AutoFit
    CapWidth ws, "B", 45
    CapWidth ws, "C", 50
    CapWidth ws, "D", 50
End Sub

'==========================================================================
' シート: 分岐カバレッジ
'==========================================================================

Private Sub RenderCoverage(ws As Worksheet, root As Object)
    ws.Cells.Clear
    Dim cov As Object
    Set cov = root("coverage")

    ws.Range("A1").Value = "分岐カバレッジ"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim rate As Double
    rate = CDbl(cov("coverageRate"))
    ws.Range("A3").Value = "カバレッジ"
    ws.Range("B3").Value = cov("coveredBranches") & " / " & cov("totalBranches") & _
                           " 分岐 (" & Format(rate, "0.0%") & ")"
    ws.Range("B3").Font.Bold = True

    Dim hdr As Long
    hdr = 5
    ws.Cells(hdr, 1).Value = "分岐ID"
    ws.Cells(hdr, 2).Value = "種別"
    ws.Cells(hdr, 3).Value = "ラベル"
    ws.Cells(hdr, 4).Value = "行番号"
    ws.Cells(hdr, 5).Value = "カバー状況"
    ws.Cells(hdr, 6).Value = "カバーするTC"
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 6)).Font.Bold = True
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 6)).Interior.Color = C_HEADER

    Dim br As Object
    Set br = cov("branches")
    Dim row As Long, pass As Long, i As Long
    row = hdr + 1
    For pass = 0 To 1
        For i = 1 To br.Count
            Dim b As Object
            Set b = br(i)
            Dim isCov As Boolean
            isCov = CBool(b("covered"))
            If (pass = 0 And Not isCov) Or (pass = 1 And isCov) Then
                ws.Cells(row, 1).Value = b("branchId")
                ws.Cells(row, 2).Value = b("type")
                ws.Cells(row, 3).Value = b("label")
                ws.Cells(row, 4).Value = b("line")
                If isCov Then
                    ws.Cells(row, 5).Value = "カバー済"
                    ws.Cells(row, 5).Interior.Color = C_OK
                Else
                    ws.Cells(row, 5).Value = "未カバー"
                    ws.Cells(row, 5).Interior.Color = C_ERR
                End If
                ws.Cells(row, 6).Value = JoinList(b("byTestCases"), ", ")
                row = row + 1
            End If
        Next i
    Next pass

    RenderWarnings ws, root, row
    ws.Columns("A:F").AutoFit
    CapWidth ws, "C", 50
    CapWidth ws, "F", 50
End Sub

'==========================================================================
' シート: 呼出関係
'==========================================================================

Private Sub RenderCallGraph(ws As Worksheet, root As Object)
    ws.Cells.Clear
    Dim cg As Object
    Set cg = root("callGraph")

    ws.Range("A1").Value = "呼出関係 (PERFORM / CALL)"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    Dim row As Long, i As Long
    row = 3
    ws.Cells(row, 1).Value = ChrW$(&H25A0) & " 段落 / SECTION 一覧"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "名称"
    ws.Cells(row, 2).Value = "種別"
    ws.Cells(row, 3).Value = "行番号"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 3)).Interior.Color = C_HEADER
    row = row + 1

    Dim nodes As Object
    Set nodes = cg("nodes")
    For i = 1 To nodes.Count
        ws.Cells(row, 1).Value = nodes(i)("name")
        ws.Cells(row, 2).Value = nodes(i)("kind")
        ws.Cells(row, 3).Value = nodes(i)("line")
        row = row + 1
    Next i

    row = row + 1
    ws.Cells(row, 1).Value = ChrW$(&H25A0) & " 呼出エッジ (呼出元ごと)"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    ws.Cells(row, 1).Value = "呼出元"
    ws.Cells(row, 3).Value = "呼出先"
    ws.Cells(row, 4).Value = "種別"
    ws.Cells(row, 5).Value = "行番号"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Font.Bold = True
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Interior.Color = C_HEADER
    row = row + 1

    Dim edges As Object
    Set edges = cg("edges")
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim j As Long, k As Long
    For j = 1 To edges.Count
        Dim fromName As String
        fromName = CStr(edges(j)("from"))
        If Not seen.Exists(fromName) Then
            seen.Add fromName, True
            ws.Cells(row, 1).Value = fromName
            ws.Cells(row, 1).Font.Bold = True
            row = row + 1
            For k = 1 To edges.Count
                If CStr(edges(k)("from")) = fromName Then
                    ws.Cells(row, 2).Value = ChrW$(&H2514) & ChrW$(&H2192)
                    ws.Cells(row, 3).Value = edges(k)("to")
                    ws.Cells(row, 4).Value = edges(k)("kind")
                    ws.Cells(row, 5).Value = edges(k)("line")
                    row = row + 1
                End If
            Next k
        End If
    Next j

    Dim ptr As Object
    Set ptr = root("programStructure")("performThruRanges")
    If ptr.Count > 0 Then
        row = row + 1
        ws.Cells(row, 1).Value = ChrW$(&H25A0) & " PERFORM THRU 範囲 (手動確認対象)"
        ws.Cells(row, 1).Font.Bold = True
        row = row + 1
        ws.Cells(row, 1).Value = "FROM"
        ws.Cells(row, 3).Value = "TO"
        ws.Cells(row, 5).Value = "行番号"
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Font.Bold = True
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 5)).Interior.Color = C_HEADER
        row = row + 1
        For i = 1 To ptr.Count
            ws.Cells(row, 1).Value = ptr(i)("from")
            ws.Cells(row, 3).Value = ptr(i)("to")
            ws.Cells(row, 5).Value = ptr(i)("line")
            row = row + 1
        Next i
    End If

    RenderWarnings ws, root, row
    ws.Columns("A:E").AutoFit
End Sub

'==========================================================================
' 共通ヘルパー
'==========================================================================

Private Sub RenderWarnings(ws As Worksheet, root As Object, ByRef row As Long)
    Dim warns As Object
    Set warns = root("warnings")
    If warns Is Nothing Then Exit Sub
    If warns.Count = 0 Then Exit Sub

    row = row + 2
    ws.Cells(row, 1).Value = ChrW$(&H25A0) & " 警告"
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1
    Dim i As Long
    For i = 1 To warns.Count
        ws.Cells(row, 1).Value = CStr(warns(i))
        ws.Cells(row, 1).Interior.Color = C_WARN
        row = row + 1
    Next i
End Sub

Private Function JoinList(lst As Object, ByVal sep As String) As String
    Dim s As String, i As Long
    s = ""
    If lst Is Nothing Then
        JoinList = ""
        Exit Function
    End If
    For i = 1 To lst.Count
        If i > 1 Then s = s & sep
        s = s & CStr(lst(i))
    Next i
    JoinList = s
End Function

Private Sub CapWidth(ws As Worksheet, ByVal col As String, ByVal maxW As Double)
    If ws.Columns(col).ColumnWidth > maxW Then ws.Columns(col).ColumnWidth = maxW
End Sub