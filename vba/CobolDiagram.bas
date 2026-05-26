Attribute VB_Name = "CobolDiagram"
' CobolDiagram - ver2.0 feature (1): draw the call / data-usage relationship
' diagram on a "呼出関係図" sheet (main program -> called sub-programs with
' their USING arguments, plus the files/DB accessed).
'
' Return-value (out) arguments cannot be auto-determined without a naming
' convention (none in this project), so each sub shows its USING arguments as
' "引数" and a "返値: 要レビュー" placeholder for human confirmation.

Option Explicit

Private Const SHP_ROUNDRECT As Long = 5    ' msoShapeRoundedRectangle
Private Const SHP_CAN As Long = 13         ' msoShapeCan (cylinder / DB look)
Private Const TXT_HORIZONTAL As Long = 1   ' msoTextOrientationHorizontal

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

    ' Draw with screen updating off (the 5-sheet report turned it back on).
    Application.ScreenUpdating = False
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo Done_

    ws.Cells.Clear
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop

    ws.Range("A1").value = "呼出 / 利用関係図 : " & progName
    With ws.Range("A1").Font
        .Name = "Meiryo UI": .Size = 14: .Bold = True: .Color = RGB(38, 70, 83)
    End With
    ws.Range("A2").value = "※ 返値(出力引数)は命名規約が無いため自動確定不可。データフロー解析+手動確認の対象です。"
    With ws.Range("A2").Font
        .Name = "Meiryo UI": .Size = 9: .Color = RGB(120, 120, 120)
    End With

    ' --- main program box (left) ---
    Dim mainLeft As Double, mainTop As Double, mainW As Double, mainH As Double
    mainLeft = 150: mainTop = 170: mainW = 140: mainH = 90
    Dim mainBox As Shape
    Set mainBox = AddBox_(ws, SHP_ROUNDRECT, mainLeft, mainTop, mainW, mainH, _
                          progName & Chr(10) & "(メイン)", RGB(255, 242, 204), RGB(40, 40, 40))

    Dim mainCx As Double, mainCy As Double
    mainCx = mainLeft + mainW
    mainCy = mainTop + mainH / 2

    ' --- data resources (cylinders, top row) ---
    Dim dx As Double, dy As Double, dw As Double, dh As Double, gi As Long
    dy = 60: dw = 140: dh = 64
    Dim d As OrderedDict, dShape As Shape, modeText As String
    gi = 0
    For Each d In data
        dx = 470 + gi * 170
        Set dShape = AddBox_(ws, SHP_CAN, dx, dy, dw, dh, CStr(d.Item("name")), RGB(218, 232, 246), RGB(40, 40, 40))
        AddConnector_ ws, mainLeft + mainW / 2, mainTop, dx + dw / 2, dy + dh
        modeText = JoinModes_(d.Item("modes"))
        AddNote_ ws, dx, dy + dh + 2, dw + 80, 16, modeText, 8, RGB(120, 120, 120)
        gi = gi + 1
    Next d

    ' --- called sub programs (right column) ---
    Dim si As Long, sx As Double, sy As Double, sw As Double, sh As Double
    sx = 470: sw = 150: sh = 60
    Dim c As OrderedDict, subBox As Shape, argText As String
    si = 0
    For Each c In calls
        sy = 170 + si * 120
        Set subBox = AddBox_(ws, SHP_ROUNDRECT, sx, sy, sw, sh, CStr(c.Item("program")), _
                             SubColor_(si), RGB(40, 40, 40))
        AddConnector_ ws, mainCx, mainCy, sx, sy + sh / 2
        argText = "引数: " & JoinArgs_(c.Item("args")) & Chr(10) & "返値: 要レビュー"
        AddNote_ ws, sx + sw + 12, sy, 240, sh, argText, 10, RGB(40, 40, 40)
        si = si + 1
    Next c

    If calls.Count = 0 And data.Count = 0 Then
        ws.Range("A4").value = "外部 CALL / ファイル・DB 利用は検出されませんでした。"
    End If

    ws.Range("A1").Select

Done_:
    Application.ScreenUpdating = True
End Sub

Private Function AddBox_(ByVal ws As Worksheet, ByVal shapeType As Long, _
                         ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
                         ByVal caption As String, ByVal fillColor As Long, ByVal textColor As Long) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(shapeType, x, y, w, h)
    s.Fill.ForeColor.RGB = fillColor
    s.Line.ForeColor.RGB = RGB(90, 90, 90)
    With s.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = caption
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Size = 11
            .Font.Bold = msoTrue
            .Font.Name = "Meiryo UI"
            .Font.Fill.ForeColor.RGB = textColor
        End With
    End With
    Set AddBox_ = s
End Function

Private Sub AddConnector_(ByVal ws As Worksheet, ByVal x1 As Double, ByVal y1 As Double, _
                          ByVal x2 As Double, ByVal y2 As Double)
    Dim ln As Shape
    Set ln = ws.Shapes.AddLine(x1, y1, x2, y2)
    ln.Line.ForeColor.RGB = RGB(120, 120, 120)
    ln.Line.Weight = 1.25
    On Error Resume Next
    ln.Line.EndArrowheadStyle = 2 ' msoArrowheadTriangle
    On Error GoTo 0
End Sub

Private Sub AddNote_(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, _
                     ByVal w As Double, ByVal h As Double, ByVal txt As String, _
                     ByVal sz As Single, ByVal col As Long)
    Dim t As Shape
    Set t = ws.Shapes.AddTextbox(TXT_HORIZONTAL, x, y, w, h)
    t.Line.Visible = msoFalse
    t.Fill.Visible = msoFalse
    With t.TextFrame2.TextRange
        .Text = txt
        .Font.Size = sz
        .Font.Name = "Meiryo UI"
        .Font.Fill.ForeColor.RGB = col
    End With
    t.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

Private Function SubColor_(ByVal idx As Long) As Long
    Select Case (idx Mod 3)
        Case 0: SubColor_ = RGB(226, 240, 217)  ' green
        Case 1: SubColor_ = RGB(218, 232, 246)  ' blue
        Case Else: SubColor_ = RGB(252, 228, 214) ' orange
    End Select
End Function

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
