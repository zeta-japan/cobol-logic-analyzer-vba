Attribute VB_Name = "Main"
' Main - entry points for the COBOL Logic Analyzer.
'   Sub_RunHello     : Phase 1 smoke test (programName + lines to A1).
'   AnalyzeAndBuild  : Phase 5 full pipeline. Reads a .cbl, runs the engine,
'                      writes a *.vba.logic.json next to the source, then
'                      calls CobolLogicViewer.BuildCobolReport to render the
'                      5 sheets in the current workbook.
'   PickCobolAndBuild: Shows a file picker, then calls AnalyzeAndBuild.

Option Explicit

Public Sub Sub_RunHello()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"

    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")

    Dim result As OrderedDict
    Set result = CobolParser.Analyze_Phase1(src)

    Dim json As String
    json = JsonWriter.WriteJson(result)

    ActiveSheet.Range("A1").value = json
    Debug.Print json
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

    Dim jsonPath As String
    jsonPath = cblPath & ".vba.logic.json"
    CobolEncoding.WriteAllText jsonPath, json, "utf-8"

    CobolLogicViewer.BuildCobolReport jsonPath
End Sub

Public Sub PickCobolAndBuild()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Select a COBOL source file"
    fd.Filters.Clear
    fd.Filters.Add "COBOL files", "*.cbl;*.cob;*.cobol;*.txt"
    fd.Filters.Add "All files", "*.*"
    fd.AllowMultiSelect = False
    If Not fd.Show Then Exit Sub
    AnalyzeAndBuild fd.SelectedItems(1)
End Sub
