Attribute VB_Name = "Main"
' Main - entry points for the COBOL Logic Analyzer.
'   SetupControlSheet: build the styled control sheet + button (run once, save).
'   PickCobolAndBuild: file picker -> AnalyzeAndBuild (the button calls this).
'   AnalyzeAndBuild  : full pipeline. Reads a .cbl, runs the engine, writes a
'                      JSON to %TEMP%, renders 5 sheets, deletes the temp JSON.
'   Sub_RunHello     : Phase 1 smoke test (programName + lines to A1).

Option Explicit

' re-entrancy latch: DoEvents heartbeats pump the message queue, and
' EnableEvents=False does NOT block button OnAction macros - without this,
' clicking 解析 again mid-run re-enters the pipeline and corrupts module
' state. TC nav buttons check it too (via AnalysisBusy).
Private mBusy As Boolean

' stage-3 status text (CobolFlow appends a live counter to it)
Public Const STATUS_FLOW As String = "解析中 (3/4): テストケース生成..."

Public Function AnalysisBusy() As Boolean
    AnalysisBusy = mBusy
End Function

' Build (or rebuild) a product-style control sheet: header band, usage card,
' a rounded "analyze" button, the output-sheet list, and a footer. Run once in
' the dev workbook, then save; everything persists in the distributed .xlsm.
Public Sub SetupControlSheet()
    Const SHEET_NAME As String = "コントロール"

    Dim cHeader As Long, cAccent As Long, cText As Long, cMuted As Long, cCard As Long, cWhite As Long
    cHeader = RGB(38, 70, 83)      ' dark slate
    cAccent = RGB(42, 157, 143)    ' teal
    cText = RGB(40, 40, 40)
    cMuted = RGB(120, 120, 120)
    cCard = RGB(244, 246, 248)
    cWhite = RGB(255, 255, 255)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SHEET_NAME
    Else
        ws.Move Before:=ThisWorkbook.Worksheets(1)
    End If

    ' reset content + shapes
    ws.Cells.Clear
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop

    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    On Error GoTo 0

    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B:H").ColumnWidth = 13.5

    ' --- header band -----------------------------------------------------
    With ws.Range("B2:H3")
        .Merge
        .Interior.Color = cHeader
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Range("B2").value = "COBOL ロジック解析ツール"
    With ws.Range("B2").Font
        .Name = "Meiryo UI"
        .Size = 20
        .Bold = True
        .Color = cWhite
    End With
    ws.Rows("2:3").RowHeight = 24

    With ws.Range("B4:H4")
        .Merge
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Range("B4").value = "COBOL ソースの分岐構造・テストケース候補・カバレッジを Excel 上で可視化します"
    With ws.Range("B4").Font
        .Name = "Meiryo UI"
        .Size = 10
        .Color = cMuted
    End With
    ws.Rows("4").RowHeight = 22

    ' --- usage card ------------------------------------------------------
    ws.Range("B6").value = "■ 使い方"
    StyleHeading_ ws.Range("B6"), cAccent
    With ws.Range("B7:H8")
        .Merge
        .Interior.Color = cCard
        .WrapText = True
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
        .Font.Name = "Meiryo UI"
        .Font.Size = 11
        .Font.Color = cText
    End With
    ws.Range("B7").value = "1.  下の［解析］ボタンを押して、解析したい COBOL ソース (.cbl / .txt) を選択" & Chr(10) & _
                           "2.  解析結果が 複数のシートに自動生成されます (再実行で自動上書き)"
    ws.Rows("7:8").RowHeight = 20

    ' --- analyze button (rounded rectangle shape) ------------------------
    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, ws.Range("B10").Left, ws.Range("B10").Top, 320, 46) ' 5=RoundedRectangle
    btn.Name = "btnAnalyze"
    btn.Fill.ForeColor.RGB = cAccent
    btn.Line.Visible = msoFalse
    On Error Resume Next
    btn.Shadow.Visible = msoFalse
    On Error GoTo 0
    With btn.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = "COBOL ソースを選択して解析"
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Size = 13
            .Font.Bold = msoTrue
            .Font.Name = "Meiryo UI"
            .Font.Fill.ForeColor.RGB = cWhite
        End With
    End With
    btn.OnAction = "PickCobolAndBuild"
    ws.Rows("10:11").RowHeight = 26

    ' --- output sheets list ----------------------------------------------
    ws.Range("B13").value = "■ 生成されるシート"
    StyleHeading_ ws.Range("B13"), cAccent
    With ws.Range("B14:H18")
        .Interior.Color = cCard
        .Font.Name = "Meiryo UI"
        .Font.Size = 10.5
        .Font.Color = cText
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B14").value = "①  COBOLソース ／ ロジック階層（ソース順 ※TC列付き・実行順展開 ※ケース標記ボタン付き）"
    ws.Range("B15").value = "②  テストケース候補（正常系=C1最少シナリオ + 異常系シナリオ、ステップフロー形式）"
    ws.Range("B16").value = "③  分岐カバレッジ表（検証Point × ケースの○マトリクス、SECTION列付き、未カバー行は赤）"
    ws.Range("B17").value = "④  入出力-想定結果（ケース毎の入力設定・出力想定値・実測記入欄）"
    ws.Range("B18").value = "⑤  Driver雛形 ／ 別紙1ドラフト（ソース＋ケースNo標記） ／ 呼出関係・呼出関係図・分岐カバレッジ"
    ws.Rows("14:18").RowHeight = 19

    ' --- footer ----------------------------------------------------------
    With ws.Range("B20:H20")
        .Merge
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
    End With
    ws.Range("B20").value = "外部依存なし (PowerShell / .NET 不要)  ・  Excel 標準機能のみ  ・  ソースは読み取り専用"
    With ws.Range("B20").Font
        .Name = "Meiryo UI"
        .Size = 9
        .Color = cMuted
    End With

    ' --- ver3.0: terminator section registration -------------------------
    ws.Range("B22").value = "■ 終了扱いセクション（ABEND処理段など・任意）"
    StyleHeading_ ws.Range("B22"), cAccent
    ws.Range("B23").value = "名前に ABEND を含む SECTION は自動で異常終了扱いになります。他の命名の終了段はここに入力してください（例: ERR-EXIT-SEC）"
    With ws.Range("B23").Font
        .Name = "Meiryo UI"
        .Size = 9
        .Color = cMuted
    End With
    With ws.Range("B24:B29")
        .Interior.Color = cWhite
        .Font.Name = "MS Gothic"
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(200, 200, 200)
        .NumberFormat = "@"
    End With

    ws.Range("A1").Select
End Sub

Private Sub StyleHeading_(ByVal rng As Range, ByVal accent As Long)
    With rng.Font
        .Name = "Meiryo UI"
        .Size = 12
        .Bold = True
        .Color = accent
    End With
End Sub

Public Sub PickCobolAndBuild()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "COBOL ソースを選択してください"
    fd.Filters.Clear
    fd.Filters.Add "COBOL files", "*.cbl;*.cob;*.cobol;*.txt"
    fd.Filters.Add "All files", "*.*"
    fd.AllowMultiSelect = False
    If Not fd.Show Then Exit Sub
    AnalyzeAndBuild fd.SelectedItems(1)
End Sub

Public Sub AnalyzeAndBuild(ByVal cblPath As String)
    If Len(Dir(cblPath)) = 0 Then
        Err.Raise 53, "Main.AnalyzeAndBuild", "File not found: " & cblPath
    End If

    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")

    If mBusy Then Exit Sub   ' re-entrant click during a running analysis
    mBusy = True

    ' performance shell: manual recalc + no events while the sheets build;
    ' restored by PerfRestore_ (HYPERLINK formulas resolve on that recalc).
    Dim prevCalc As Long
    prevCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.StatusBar = "解析中 (1/4): エンジン解析・パス列挙..."

    Dim result As OrderedDict, engineErr As String, engineErrN As Long
    Dim json As String
    On Error Resume Next
    Set result = CobolParser.Analyze_Full(src, "", "utf-8")
    If Err.Number = 0 And Not result Is Nothing Then json = JsonWriter.WriteJson(result)
    If Err.Number <> 0 Then
        engineErrN = Err.Number
        engineErr = "#" & Err.Number & " " & Err.Description
    End If
    On Error GoTo 0
    If result Is Nothing Or engineErrN <> 0 Then
        PerfRestore_ prevCalc
        Err.Raise 5, "Main.AnalyzeAndBuild", "エンジン解析に失敗しました " & engineErr
    End If

    ' Write the intermediate JSON to TEMP (never next to the source), hand it to
    ' the viewer, then delete it. The viewer reads the whole file into memory up
    ' front, so nothing accumulates on disk and the user never has to clean up.
    Dim baseName As String
    baseName = Mid$(cblPath, InStrRev(cblPath, "\") + 1)
    Dim jsonPath As String
    jsonPath = Environ$("TEMP") & "\CobolAnalyzer_" & baseName & ".logic.json"
    On Error Resume Next
    CobolEncoding.WriteAllText jsonPath, json, "utf-8"
    If Err.Number <> 0 Then
        engineErrN = Err.Number
        engineErr = "#" & Err.Number & " " & Err.Description
        On Error GoTo 0
        PerfRestore_ prevCalc
        Err.Raise 5, "Main.AnalyzeAndBuild", "中間ファイル出力に失敗しました " & engineErr
    End If
    On Error GoTo 0

    ' ver2.2: the hierarchy sheet was split into (ソース順)/(実行順展開).
    ' Drop the legacy combined sheet so an old copy cannot linger stale.
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("ロジック階層").Delete
    ThisWorkbook.Sheets("入力項目").Delete
    ThisWorkbook.Sheets("Driver_Dummy雛形").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Suppress this call's own "done" dialog; one notice is shown at the very end
    ' so it does not interrupt before the ver2.0 sheets are built.
    Application.StatusBar = "解析中 (2/4): 基本シート描画..."
    CobolLogicViewer.BuildCobolReport jsonPath, False

    ' ver3.0: execution-flow case generation -> tree marking + case sheets
    Application.StatusBar = STATUS_FLOW
    Dim flowR As OrderedDict, flowErr As String
    flowErr = ""
    On Error Resume Next
    Set flowR = CobolFlow.Analyze_Flow(src, Get_TermSections())
    If Err.Number <> 0 Then flowErr = "#" & Err.Number & " " & Err.Description
    On Error GoTo 0
    ' BuildTcMarking must run even when flowR is Nothing - it clears the
    ' previous run's nav buttons / scratch data (stale-marking guard).
    On Error Resume Next
    CobolTcMark.BuildTcMarking flowR
    ' same stale-guard contract: clears the helper-token column even on failure
    CobolXdm.ApplyTreeTc flowR
    On Error GoTo 0
    If Not flowR Is Nothing Then
        On Error Resume Next
        CobolCaseView.BuildCaseSheets flowR
        CobolXdm.BuildBesshiDraft flowR, src
        On Error GoTo 0
    Else
        FlowFailBanner_ flowErr
    End If

    Application.StatusBar = "解析中 (4/4): 付帯シート生成..."
    ' ver2.0 feature (1): call/usage relationship diagram sheet
    On Error Resume Next
    CobolDiagram.BuildCallDiagram cblPath
    On Error GoTo 0

    ' ver3.0 P4: per-case input setup / expected results sheet
    If Not flowR Is Nothing Then
        On Error Resume Next
        CobolIoView.BuildIoSheet flowR, src
        On Error GoTo 0
    End If

    ' ver3.0 P5: per-case Driver skeleton (Dummy retired)
    If Not flowR Is Nothing Then
        On Error Resume Next
        CobolStub.BuildDriverSheet flowR, cblPath
        On Error GoTo 0
    End If

    ' reference sheets the team rarely opens go to the end of the tab order
    On Error Resume Next
    Dim shMove As Variant
    For Each shMove In Array("分岐カバレッジ", "呼出関係", "呼出関係図")
        ThisWorkbook.Sheets(CStr(shMove)).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Next shMove
    On Error GoTo 0

    On Error Resume Next
    Kill jsonPath
    On Error GoTo 0

    PerfRestore_ prevCalc

    ' Single completion notice, after every sheet (incl. ver2.0) is built.
    On Error Resume Next
    ThisWorkbook.Sheets("ロジック階層(実行順展開)").Activate
    On Error GoTo 0
    MsgBox "解析が完了しました。" & vbLf & vbLf & _
           "生成シート:" & vbLf & _
           "・COBOLソース / ロジック階層(ソース順) / ロジック階層(実行順展開)" & vbLf & _
           "・テストケース候補 / 分岐カバレッジ表 / 分岐カバレッジ" & vbLf & _
           "・呼出関係 / 呼出関係図 / 入出力-想定結果 / Driver雛形", vbInformation
End Sub

' ver3.0: terminator sections registered on the control sheet (B24:B29).
' A path reaching PERFORM <one of these> is treated as an abnormal end
' (ABEND handler etc.). Empty/missing sheet = no registered terminators.
' Restore the performance shell (one recalc resolves HYPERLINK formulas).
Private Sub PerfRestore_(ByVal prevCalc As Long)
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculate   ' resolve HYPERLINK formulas even if the user runs manual calc
    Application.Calculation = prevCalc
    mBusy = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' ver3.0: when case generation fails, show a clear banner on the case sheets
' (instead of silently leaving the previous content) with the engine error.
Private Sub FlowFailBanner_(ByVal msg As String)
    On Error Resume Next
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array("テストケース候補", "分岐カバレッジ表", "入出力-想定結果", "Driver雛形", "別紙1ドラフト")
        Set ws = JsonParser.EnsureSheet(CStr(nm))
        ws.Cells.Clear
        ws.Range("A1").value = "ver3.0 ケース生成に失敗しました（解析エラー）"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Color = RGB(192, 0, 0)
        ws.Range("A2").value = "エラー: " & msg & "  → この表示のスクリーンショットを開発側へ共有してください"
    Next nm
    On Error GoTo 0
End Sub
Public Function Get_TermSections() As Collection
    Dim c As Collection
    Set c = New Collection
    Set Get_TermSections = c
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("コントロール")
    If ws Is Nothing Then Exit Function
    Dim r As Long, v As String
    For r = 24 To 29
        v = Trim$(CStr(ws.Cells(r, 2).value))
        If Len(v) > 0 Then c.Add UCase$(v)
    Next r
    On Error GoTo 0
End Function

' Phase 1 smoke test: write a minimal JSON summary to A1 of the active sheet.
Public Sub Sub_RunHello()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"

    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")

    Dim result As OrderedDict
    Set result = CobolParser.Analyze_Phase1(src)

    ActiveSheet.Range("A1").value = JsonWriter.WriteJson(result)
End Sub
