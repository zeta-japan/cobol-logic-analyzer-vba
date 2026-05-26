Attribute VB_Name = "Main"
' Main - entry points for the COBOL Logic Analyzer.
'   SetupControlSheet: build the styled control sheet + button (run once, save).
'   PickCobolAndBuild: file picker -> AnalyzeAndBuild (the button calls this).
'   AnalyzeAndBuild  : full pipeline. Reads a .cbl, runs the engine, writes a
'                      JSON to %TEMP%, renders 5 sheets, deletes the temp JSON.
'   Sub_RunHello     : Phase 1 smoke test (programName + lines to A1).

Option Explicit

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
    ws.Range("B7").value = "1.  下の［解析］ボタンを押して、解析したい COBOL ソース (.cbl) を選択" & Chr(10) & _
                           "2.  解析結果が 5 つのシートに自動生成されます (再実行で自動上書き)"
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
    ws.Range("B14").value = "①  COBOLソース       元コード + 通し行番号"
    ws.Range("B15").value = "②  ロジック階層       IF / EVALUATE / SEARCH の入れ子をツリー表示"
    ws.Range("B16").value = "③  テストケース候補   全実行パスからテストケース表を自動生成"
    ws.Range("B17").value = "④  分岐カバレッジ     各分岐の網羅状況と警告"
    ws.Range("B18").value = "⑤  呼出関係           PERFORM / CALL / 段落 の呼出関係"
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

    Dim result As OrderedDict
    Set result = CobolParser.Analyze_Full(src, "", "utf-8")

    Dim json As String
    json = JsonWriter.WriteJson(result)

    ' Write the intermediate JSON to TEMP (never next to the source), hand it to
    ' the viewer, then delete it. The viewer reads the whole file into memory up
    ' front, so nothing accumulates on disk and the user never has to clean up.
    Dim baseName As String
    baseName = Mid$(cblPath, InStrRev(cblPath, "\") + 1)
    Dim jsonPath As String
    jsonPath = Environ$("TEMP") & "\CobolAnalyzer_" & baseName & ".logic.json"
    CobolEncoding.WriteAllText jsonPath, json, "utf-8"

    CobolLogicViewer.BuildCobolReport jsonPath

    On Error Resume Next
    Kill jsonPath
    On Error GoTo 0
End Sub

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
