Attribute VB_Name = "CobolDataView"
' CobolDataView - ver2.0 feature (2): render the data items / arguments table
' on a "入力項目" sheet. LINKAGE items (the program's external arguments) are
' highlighted; COPY lines are shown greyed as "未展開".

Option Explicit

' retired in ver3.0 - replaced by CobolIoView (kept for reference)
Private Sub BuildDataItemsSheet(ByVal cblPath As String)
    Dim src As String
    src = CobolEncoding.ReadCobolSource(cblPath, "auto")
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim items As Collection
    Set items = CobolData.Get_DataItems(norm.Item("Lines"))

    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet("入力項目")

    Application.ScreenUpdating = False
    On Error GoTo Done_
    ws.Cells.Clear

    ws.Range("A1").value = "入力項目 / 引数一覧"
    With ws.Range("A1").Font
        .Name = "Meiryo UI": .Size = 14: .Bold = True: .Color = RGB(38, 70, 83)
    End With
    ws.Range("A2").value = "※ LINKAGE SECTION の項目が外部引数(黄色)。COPY句は copybook 本体が無いため未展開です。"
    With ws.Range("A2").Font
        .Name = "Meiryo UI": .Size = 9: .Color = RGB(120, 120, 120)
    End With

    Dim hdr As Long
    hdr = 4
    Dim headers As Variant
    headers = Array("区分", "レベル", "項目名", "PIC", "型", "桁", "OCCURS", "初期値", "値候補")
    Dim cI As Long
    For cI = 0 To UBound(headers)
        ws.Cells(hdr, cI + 1).value = headers(cI)
    Next cI
    With ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 9))
        .Font.Bold = True
        .Font.Name = "Meiryo UI"
        .Interior.Color = RGB(217, 225, 232)
    End With

    Dim n As Long
    n = items.Count
    If n = 0 Then
        ws.Cells(hdr + 1, 1).value = "データ項目は検出されませんでした。"
        GoTo Done_
    End If

    Dim arr() As Variant
    ReDim arr(1 To n, 1 To 9)
    Dim it As OrderedDict, r As Long
    r = 1
    For Each it In items
        arr(r, 1) = it.Item("section")
        arr(r, 2) = it.Item("level")
        arr(r, 3) = it.Item("name")
        arr(r, 4) = it.Item("pic")
        arr(r, 5) = TypeLabel_(CStr(it.Item("picType")))
        If CLng(it.Item("picLen")) > 0 Then arr(r, 6) = CLng(it.Item("picLen")) Else arr(r, 6) = ""
        If CStr(it.Item("occurs")) <> "" Then arr(r, 7) = "x" & it.Item("occurs") Else arr(r, 7) = ""
        arr(r, 8) = it.Item("value")
        If it.Item("isCopy") Then
            arr(r, 9) = ""
        Else
            arr(r, 9) = CobolData.ValueHint_(CStr(it.Item("picType")), CLng(it.Item("picLen")))
        End If
        r = r + 1
    Next it
    ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(hdr + n, 9)).value = arr
    ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(hdr + n, 9)).Font.Name = "Meiryo UI"

    ' highlight LINKAGE rows (arguments); grey COPY rows
    r = 1
    For Each it In items
        If UCase$(CStr(it.Item("section"))) = "LINKAGE" Then
            ws.Range(ws.Cells(hdr + r, 1), ws.Cells(hdr + r, 9)).Interior.Color = RGB(255, 247, 214)
        ElseIf it.Item("isCopy") Then
            ws.Range(ws.Cells(hdr + r, 1), ws.Cells(hdr + r, 9)).Interior.Color = RGB(238, 238, 238)
        End If
        r = r + 1
    Next it

    ws.Columns("A:I").AutoFit
    ws.Range("A1").Select
Done_:
    Application.ScreenUpdating = True
End Sub

Private Function TypeLabel_(ByVal t As String) As String
    Select Case t
        Case "num": TypeLabel_ = "数値"
        Case "signed-num": TypeLabel_ = "符号付数値"
        Case "decimal": TypeLabel_ = "小数"
        Case "alnum": TypeLabel_ = "英数字"
        Case "group": TypeLabel_ = "集団項目"
        Case "copy": TypeLabel_ = "COPY(未展開)"
        Case Else: TypeLabel_ = t
    End Select
End Function
