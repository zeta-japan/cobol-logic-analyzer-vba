Attribute VB_Name = "Main"
' Main - entry points for the COBOL Logic Analyzer.
'   SetupControlSheet: build the control sheet + button (run once, then save).
'   PickCobolAndBuild: file picker -> AnalyzeAndBuild (the button calls this).
'   AnalyzeAndBuild  : full pipeline. Reads a .cbl, runs the engine, writes a
'                      JSON to %TEMP%, renders 5 sheets, deletes the temp JSON.
'   Sub_RunHello     : Phase 1 smoke test (programName + lines to A1).

Option Explicit

' Build (or rebuild) the control sheet with a single button that runs the
' analysis. Run this once in the dev workbook, then save; the button and sheet
' persist, so end users only ever click the button.
Public Sub SetupControlSheet()
    Const SHEET_NAME As String = "コントロール"
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

    ws.Cells.Clear
    Dim shp As Object
    For Each shp In ws.Buttons: shp.Delete: Next shp

    ws.Range("A1").value = "COBOL ロジック解析ツール"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A3").value = "■ 使い方"
    ws.Range("A3").Font.Bold = True
    ws.Range("A4").value = "   1. 下のボタンを押して COBOL ソース (.cbl) を選択"
    ws.Range("A5").value = "   2. 解析結果が 5 シートに自動生成されます"
    ws.Range("A6").value = "      (COBOLソース / ロジック階層 / テストケース候補 / 分岐カバレッジ / 呼出関係)"

    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("A8").Left, ws.Range("A8").Top, 240, 36)
    btn.Caption = "COBOL ソースを選択して解析"
    btn.Name = "btnAnalyze"
    btn.OnAction = "PickCobolAndBuild"

    ws.Activate
    ws.Range("A1").Select
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
