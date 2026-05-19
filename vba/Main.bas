Attribute VB_Name = "Main"
' Main - entry points for the Pure VBA COBOL Logic Analyzer.
' Phase 1: Sub_RunHello reads samples/input/ICASE1.cbl relative to the
' workbook, runs Analyze_Phase1, and writes the JSON to cell A1 of the
' active sheet. Later phases add AnalyzeAndBuild for full 5-sheet rendering.

Option Explicit

Public Sub Sub_RunHello()
    Dim cblPath As String
    cblPath = ThisWorkbook.path & "\samples\input\ICASE1.cbl"

    Dim src As String
    src = CobolEncoding.ReadAllText(cblPath, "auto")

    Dim result As OrderedDict
    Set result = CobolParser.Analyze_Phase1(src)

    Dim json As String
    json = JsonWriter.WriteJson(result)

    ActiveSheet.Range("A1").value = json
    Debug.Print json
End Sub
