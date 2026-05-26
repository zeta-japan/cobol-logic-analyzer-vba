Attribute VB_Name = "CobolDiagram"
' CobolDiagram - ver2.0 feature (1): call / data-usage relationship on a
' "呼出関係図" sheet, rendered as a cell-based tree in the same style as the
' ロジック階層 sheet (box-drawing connectors, MS Gothic), with columns for
' 種別 / 引数(またはアクセス) / 返値.
'
' Children of the main program = data resources (DB/files) + called sub
' programs. Return values cannot be auto-determined (no naming convention),
' so calls show "要レビュー".

Option Explicit

Public Sub BuildCallDiagram(ByVal cblPath As String)
    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")

    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim progName As String
    progName = CobolParser.Get_ProgramName(lines)
    Dim calls As Collection, data As Collection
    Set calls = CobolCalls.Get_ExternalCalls(lines)
    Set data = CobolCalls.Get_DataAccess(lines)

    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet("呼出関係図")

    Application.ScreenUpdating = False
    On Error GoTo Done_
    ws.Cells.Clear
    ' drop any shapes left by the previous (shape-based) version
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Done_

    ws.Range("A1").value = "呼出 / 利用関係図 : " & progName
    With ws.Range("A1").Font
        .Name = "Meiryo UI": .Size = 14: .Bold = True: .Color = RGB(38, 70, 83)
    End With
    ws.Range("A2").value = "※ 返値(出力引数)は命名規約が無いため自動確定不可。データフロー解析+手動確認の対象です。"
    With ws.Range("A2").Font
        .Name = "Meiryo UI": .Size = 9: .Color = RGB(120, 120, 120)
    End With

    Dim hdr As Long
    hdr = 4
    ws.Cells(hdr, 1).value = "呼出 / 利用ツリー"
    ws.Cells(hdr, 2).value = "種別"
    ws.Cells(hdr, 3).value = "引数 / アクセス"
    ws.Cells(hdr, 4).value = "返値"
    With ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 4))
        .Font.Bold = True
        .Font.Name = "Meiryo UI"
        .Interior.Color = RGB(217, 225, 232)
    End With

    Dim row As Long
    row = hdr + 1

    ' root (main program)
    ws.Cells(row, 1).value = progName & "  (メイン)"
    ws.Cells(row, 1).Font.Bold = True
    ws.Cells(row, 1).Font.Name = "MS Gothic"
    ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Interior.Color = RGB(255, 242, 204)
    row = row + 1

    Dim total As Long, idx As Long
    total = data.Count + calls.Count
    idx = 0

    ' children: data resources first (DB / files)
    Dim d As OrderedDict
    For Each d In data
        idx = idx + 1
        ws.Cells(row, 1).value = TreeConnector_(idx, total) & CStr(d.Item("name"))
        ws.Cells(row, 1).Font.Name = "MS Gothic"
        ws.Cells(row, 2).value = "DB/ファイル"
        ws.Cells(row, 3).value = JoinModes_(d.Item("modes"))
        ws.Range(ws.Cells(row, 1), ws.Cells(row, 4)).Interior.Color = RGB(218, 232, 246)
        row = row + 1
    Next d

    ' children: called sub programs
    Dim c As OrderedDict
    For Each c In calls
        idx = idx + 1
        ws.Cells(row, 1).value = TreeConnector_(idx, total) & CStr(c.Item("program"))
        ws.Cells(row, 1).Font.Name = "MS Gothic"
        ws.Cells(row, 2).value = "CALL"
        ws.Cells(row, 3).value = JoinArgs_(c.Item("args"))
        ws.Cells(row, 4).value = "要レビュー"
        row = row + 1
    Next c

    If total = 0 Then
        ws.Cells(row, 1).value = "外部 CALL / ファイル・DB 利用は検出されませんでした。"
    End If

    ws.Columns("A:D").AutoFit
    CapColWidth_ ws, "A", 40
    CapColWidth_ ws, "C", 55
    ws.Range("A1").Select
Done_:
    Application.ScreenUpdating = True
End Sub

' └──  for the last child, ├──  otherwise (same glyphs as the ロジック階層 tree).
Private Function TreeConnector_(ByVal idx As Long, ByVal total As Long) As String
    If idx = total Then
        TreeConnector_ = ChrW$(&H2514) & ChrW$(&H2500) & ChrW$(&H2500) & " "
    Else
        TreeConnector_ = ChrW$(&H251C) & ChrW$(&H2500) & ChrW$(&H2500) & " "
    End If
End Function

Private Sub CapColWidth_(ByVal ws As Worksheet, ByVal col As String, ByVal maxW As Double)
    If ws.Columns(col).ColumnWidth > maxW Then ws.Columns(col).ColumnWidth = maxW
End Sub

Private Function JoinArgs_(ByVal args As Collection) As String
    Dim s As String, v As Variant, first As Boolean
    first = True
    For Each v In args
        If Not first Then s = s & ", "
        s = s & CStr(v)
        first = False
    Next v
    If s = "" Then s = "(なし)"
    JoinArgs_ = s
End Function

Private Function JoinModes_(ByVal modes As Collection) As String
    Dim s As String, v As Variant, first As Boolean
    first = True
    For Each v In modes
        If Not first Then s = s & " / "
        s = s & CStr(v)
        first = False
    Next v
    JoinModes_ = s
End Function
