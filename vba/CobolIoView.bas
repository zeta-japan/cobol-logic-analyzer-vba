Attribute VB_Name = "CobolIoView"
' CobolIoView - ver3.0 P4: 入出力-想定結果 sheet (replaces the old 入力項目).
' For every NORMAL case, derives along its execution path:
'   inputs  : (1) LINKAGE input items used on the path
'             (2) DB prerequisites (EXEC + the ADACODE arm taken)
'             (3) subprogram-return prerequisites (CALL + the arm on its output)
'   outputs : last-write-wins assignments to LINKAGE items, value shown as
'             literal / =入力X / sub-return / DB string / computed expression
'   refs    : last values of internal key items (参考)
' BuildIoModel is pure logic (shared with the Driver generator in CobolStub);
' BuildIoSheet renders the blocks with 実測値/判定 fill-in columns.

Option Explicit

Private Const SHEET_IO As String = "入出力-想定結果"

'======================================================================
' Pure model (consumed by this sheet AND the Driver generator)
'======================================================================
' Returns a Collection of OrderedDict per NORMAL case:
'   id, kindSerial, termJp,
'   lkIn  : Collection of {Item, Note}
'   dbPre : Collection of String
'   subPre: Collection of String
'   outs  : Collection of {Item, Val}
'   refs  : Collection of {Item, Val}
Public Function BuildIoModel(ByVal flow As OrderedDict, ByVal src As String) As Collection
    Dim model As Collection
    Set model = New Collection
    Set BuildIoModel = model
    If flow Is Nothing Then Exit Function
    If Not flow.Exists("cases") Then Exit Function

    ' linkage membership: explicit LINKAGE items + "01name-" prefixes (COPY
    ' PREFIXING children are not expanded, so prefix matching covers them)
    Dim norm As OrderedDict
    Set norm = CobolParser.Get_NormalizedCobolLines(src, "")
    Dim items As Collection
    Set items = CobolData.Get_DataItems(norm.Item("Lines"))
    Dim lkSet As OrderedDict, lk01 As Collection
    Set lkSet = New OrderedDict
    Set lk01 = New Collection
    Dim it As OrderedDict
    For Each it In items
        If CStr(it.Item("section")) = "LINKAGE" Then
            If Not lkSet.Exists(UCase$(CStr(it.Item("name")))) Then lkSet.Add UCase$(CStr(it.Item("name"))), True
            If CStr(it.Item("level")) = "01" Or CStr(it.Item("level")) = "1" Then lk01.Add UCase$(CStr(it.Item("name")))
        End If
    Next it

    Dim descMap As OrderedDict
    Set descMap = Nothing
    If flow.Exists("descMap") Then Set descMap = flow.Item("descMap")

    Dim c As OrderedDict
    For Each c In flow.Item("cases")
        If CStr(c.Item("kind")) = "normal" Then
            model.Add CaseIo_(c, lkSet, lk01, descMap)
        End If
    Next c
End Function

Private Function CaseIo_(ByVal c As OrderedDict, ByVal lkSet As OrderedDict, ByVal lk01 As Collection, _
                         ByVal descMap As OrderedDict) As OrderedDict
    Dim lkIn As Collection, dbPre As Collection, subPre As Collection
    Dim outs As OrderedDict, refs As OrderedDict
    Set lkIn = New Collection
    Set dbPre = New Collection
    Set subPre = New Collection
    Set outs = New OrderedDict    ' item -> value text (last write wins)
    Set refs = New OrderedDict

    ' writers: item -> {Kind: literal/move/call/string/compute/init, Src, Target}
    Dim writers As OrderedDict
    Set writers = New OrderedDict
    Dim lastExec As String
    lastExec = ""

    Dim seenLk As OrderedDict
    Set seenLk = New OrderedDict

    Dim e As OrderedDict, k As String
    For Each e In c.Item("events")
        k = CStr(e.Item("Kind"))
        If k = "exec" Then
            lastExec = Mid$(CStr(e.Item("Text")), 6)   ' drop "EXEC "
            If Left$(lastExec, 7) = "ADABAS " Then lastExec = Mid$(lastExec, 8)
        ElseIf k = "call" Then
            MarkCallWriters_ writers, e, descMap
        ElseIf k = "assign" Then
            HandleAssign_ e, writers, lkSet, lk01, seenLk, lkIn, outs, refs
        ElseIf k = "arm" Then
            HandleArm_ e, writers, lkSet, lk01, seenLk, lkIn, dbPre, subPre, lastExec
        End If
    Next e

    Dim r As OrderedDict
    Set r = New OrderedDict
    r.Add "id", c.Item("id")
    r.Add "kindSerial", c.Item("kindSerial")
    r.Add "termJp", TermJp_(c)
    r.Add "lkIn", lkIn
    r.Add "dbPre", dbPre
    r.Add "subPre", subPre
    r.Add "outs", PairList_(outs)
    r.Add "refs", PairList_(refs)
    Set CaseIo_ = r
End Function

Private Sub MarkCallWriters_(ByVal writers As OrderedDict, ByVal e As OrderedDict, ByVal descMap As OrderedDict)
    ' everything passed via USING may be (re)written by the callee
    Static rxCall As Object
    If rxCall Is Nothing Then
        Set rxCall = CreateObject("VBScript.RegExp")
        rxCall.Pattern = "^CALL\s+'([A-Z0-9-]+)'(\s+USING\s+(.+))?$"
        rxCall.IgnoreCase = False
    End If
    Dim m As Object
    Set m = rxCall.Execute(CStr(e.Item("Text")))
    If m.Count = 0 Then Exit Sub
    Dim tgt As String, params As String
    tgt = m.Item(0).SubMatches(0)
    params = m.Item(0).SubMatches(2)
    If Len(params) = 0 Then Exit Sub
    Dim pa() As String, i As Long, w As OrderedDict, v As Variant
    pa = Split(Trim$(params), " ")
    For i = LBound(pa) To UBound(pa)
        If Len(Trim$(pa(i))) > 0 Then
            Set w = New OrderedDict
            w.Add "Kind", "call"
            w.Add "Target", tgt
            writers.Add UCase$(Trim$(pa(i))), w
            If Not descMap Is Nothing Then
                If descMap.Exists(UCase$(Trim$(pa(i)))) Then
                    For Each v In descMap.Item(UCase$(Trim$(pa(i))))
                        writers.Add CStr(v), w
                    Next v
                End If
            End If
        End If
    Next i
End Sub

Private Sub HandleAssign_(ByVal e As OrderedDict, ByVal writers As OrderedDict, _
                          ByVal lkSet As OrderedDict, ByVal lk01 As Collection, ByVal seenLk As OrderedDict, _
                          ByVal lkIn As Collection, ByVal outs As OrderedDict, ByVal refs As OrderedDict)
    Dim dst As String, srcT As String, ak As String
    dst = CStr(e.Item("Dst"))
    srcT = CStr(e.Item("Src"))
    ak = CStr(e.Item("AKind"))

    ' linkage item used as a SOURCE = an input the tester must set
    If ak = "move" Then
        Dim st As String
        st = UCase$(Trim$(srcT))
        If IsLinkage_(st, lkSet, lk01) Then
            NoteLkInput_ lkIn, seenLk, st, "→ " & dst & " へ転記"
        End If
    End If

    ' record writer
    Dim w As OrderedDict
    Set w = New OrderedDict
    w.Add "Kind", ak
    w.Add "Src", srcT
    w.Add "Target", ""
    writers.Add dst, w

    ' expected output (last write wins) for linkage targets
    If IsLinkage_(dst, lkSet, lk01) Then
        outs.Add dst, ValText_(srcT, ak, writers, lkSet, lk01)
    ElseIf CBool(e.Item("IsKey")) Then
        refs.Add dst, ValText_(srcT, ak, writers, lkSet, lk01)
    End If
End Sub

Private Sub HandleArm_(ByVal e As OrderedDict, ByVal writers As OrderedDict, _
                       ByVal lkSet As OrderedDict, ByVal lk01 As Collection, ByVal seenLk As OrderedDict, _
                       ByVal lkIn As Collection, ByVal dbPre As Collection, ByVal subPre As Collection, _
                       ByVal lastExec As String)
    Dim cond As String, pol As String
    cond = CStr(e.Item("Cond"))
    pol = Polarity_(CStr(e.Item("Arm")), cond)

    Dim ids As Collection, v As Variant, done As Boolean
    Set ids = IdentsOf_(cond)
    done = False
    For Each v In ids
        If CStr(v) = "ADACODE" Then
            If Len(lastExec) > 0 Then
                dbPre.Add "「" & lastExec & "」 → " & pol
            Else
                dbPre.Add pol
            End If
            done = True
            Exit For
        End If
    Next v
    If done Then Exit Sub

    For Each v In ids
        If writers.Exists(CStr(v)) Then
            If CStr(writers.Item(CStr(v)).Item("Kind")) = "call" Then
                subPre.Add CStr(writers.Item(CStr(v)).Item("Target")) & " 戻り値: " & pol
                done = True
                Exit For
            End If
        End If
    Next v
    If done Then Exit Sub

    ' a linkage item compared directly = an input condition
    For Each v In ids
        If IsLinkage_(CStr(v), lkSet, lk01) Then
            NoteLkInput_ lkIn, seenLk, CStr(v), pol
            done = True
        End If
    Next v
End Sub

Private Function ValText_(ByVal srcT As String, ByVal ak As String, ByVal writers As OrderedDict, _
                          ByVal lkSet As OrderedDict, ByVal lk01 As Collection) As String
    Dim lit As String
    Select Case ak
        Case "move"
            If IsLiteral_(srcT, lit) Then
                ValText_ = srcT
            ElseIf IsLinkage_(UCase$(Trim$(srcT)), lkSet, lk01) Then
                ValText_ = "＝入力 " & srcT
            ElseIf writers.Exists(UCase$(Trim$(srcT))) Then
                If CStr(writers.Item(UCase$(Trim$(srcT))).Item("Kind")) = "call" Then
                    ValText_ = CStr(writers.Item(UCase$(Trim$(srcT))).Item("Target")) & " 戻り値（" & srcT & "）"
                Else
                    ValText_ = "＝ " & srcT & "（内部項目）"
                End If
            Else
                ValText_ = "＝ " & srcT & "（内部項目・要確認）"
            End If
        Case "string"
            ValText_ = "連結結果（" & srcT & "）"
        Case "compute"
            ValText_ = "計算結果（" & srcT & "）"
        Case "init"
            ValText_ = "初期化（INITIALIZE）"
        Case "accept"
            ValText_ = "ACCEPT 入力値"
        Case Else
            ValText_ = srcT
    End Select
End Function

Private Sub NoteLkInput_(ByVal lkIn As Collection, ByVal seenLk As OrderedDict, _
                         ByVal item As String, ByVal note As String)
    If seenLk.Exists(item) Then Exit Sub
    seenLk.Add item, True
    Dim r As OrderedDict
    Set r = New OrderedDict
    r.Add "Item", item
    r.Add "Note", note
    lkIn.Add r
End Sub

Private Function Polarity_(ByVal arm As String, ByVal cond As String) As String
    Select Case arm
        Case "THEN", "WHEN", "AT END"
            Polarity_ = "「" & cond & "」が成立"
        Case "ELSE"
            Polarity_ = "「" & cond & "」が不成立"
        Case Else
            Polarity_ = "「" & cond & "」（" & arm & "）"
    End Select
End Function

Private Function IsLinkage_(ByVal nm As String, ByVal lkSet As OrderedDict, ByVal lk01 As Collection) As Boolean
    IsLinkage_ = False
    If lkSet.Exists(nm) Then
        IsLinkage_ = True
        Exit Function
    End If
    Dim v As Variant
    For Each v In lk01
        If Left$(nm, Len(CStr(v)) + 1) = CStr(v) & "-" Then
            IsLinkage_ = True
            Exit Function
        End If
    Next v
End Function

Private Function IsLiteral_(ByVal s As String, ByRef out As String) As Boolean
    Dim x As String
    x = Trim$(s)
    IsLiteral_ = False
    If Len(x) >= 2 Then
        If Left$(x, 1) = "'" And Right$(x, 1) = "'" Then
            out = x
            IsLiteral_ = True
            Exit Function
        End If
    End If
    If x Like String(Len(x), "#") And Len(x) > 0 Then
        out = x
        IsLiteral_ = True
        Exit Function
    End If
    Select Case x
        Case "ZERO", "ZEROS", "ZEROES", "SPACE", "SPACES"
            out = x
            IsLiteral_ = True
    End Select
End Function

Private Function IdentsOf_(ByVal cond As String) As Collection
    Static rxId As Object
    If rxId Is Nothing Then
        Set rxId = CreateObject("VBScript.RegExp")
        rxId.Pattern = "[A-Z0-9][A-Z0-9-]+"
        rxId.Global = True
        rxId.IgnoreCase = False
    End If
    Dim c As Collection
    Set c = New Collection
    Dim m As Object, i As Long, w As String
    Set m = rxId.Execute(cond)
    For i = 0 To m.Count - 1
        w = m.Item(i).Value
        Select Case w
            Case "OR", "AND", "NOT", "ZERO", "ZEROS", "ZEROES", "SPACE", "SPACES", "OTHER", "THEN"
            Case Else
                If Not w Like String(Len(w), "#") Then c.Add w
        End Select
    Next i
    Set IdentsOf_ = c
End Function

Private Function PairList_(ByVal d As OrderedDict) As Collection
    Dim c As Collection, ks As Collection, v As Variant, r As OrderedDict
    Set c = New Collection
    Set ks = d.Keys
    For Each v In ks
        Set r = New OrderedDict
        r.Add "Item", CStr(v)
        r.Add "Val", CStr(d.Item(CStr(v)))
        c.Add r
    Next v
    Set PairList_ = c
End Function

Private Function TermJp_(ByVal c As OrderedDict) As String
    Dim t As String, tv As String
    t = CStr(c.Item("term"))
    If t = "goback" Then
        If c.Exists("termVerb") Then tv = CStr(c.Item("termVerb"))
        If Len(tv) = 0 Then tv = "GOBACK"
        TermJp_ = tv & "（正常終了）"
    ElseIf Left$(t, 6) = "abend:" Then
        TermJp_ = "異常終了（" & Mid$(t, 7) & " 経由）"
    ElseIf Left$(t, 6) = "synth:" Then
        TermJp_ = "呼出先異常（CALL " & Mid$(t, 7) & "・合成）"
    Else
        TermJp_ = t
    End If
End Function

'======================================================================
' Sheet rendering
'======================================================================
Public Sub BuildIoSheet(ByVal flow As OrderedDict, ByVal src As String)
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = JsonParser.EnsureSheet(SHEET_IO)
    ws.Cells.Clear
    ws.Columns("A:E").NumberFormat = "@"
    On Error GoTo Fail_

    Dim model As Collection
    Set model = BuildIoModel(flow, src)

    ' title band: dark navy with white text (JP corporate sheet style)
    ws.Range("A1").Value = "入出力-想定結果（正常系ケース毎の入力設定と出力想定値 ／ Driver雛形・テストケース候補と対応）"
    With ws.Range("A1")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2").Value = "※ 設定値が静的に決まらない項目は条件を表示しています。実測値・判定はテスト時に記入してください。"
    ws.Range("A2").Font.Color = RGB(120, 120, 120)
    ws.Range("A2").Font.Size = 9

    Dim row As Long
    row = 4
    Dim cm As OrderedDict
    For Each cm In model
        row = RenderCaseIo_(ws, cm, row) + 1
    Next cm

    If model.Count = 0 Then
        ws.Cells(row, 1).Value = "（正常系ケースがありません）"
    End If

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C").ColumnWidth = 56
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 10
    GoTo Done_
Fail_:
    Dim eN As Long, eD As String
    eN = Err.Number
    eD = Err.Description
    On Error Resume Next
    ws.Cells(1, 1).Interior.ColorIndex = xlNone   ' clear the navy band for readable red text
    ws.Cells(1, 1).Value = "(診断: 入出力-想定結果の生成でエラー #" & eN & " " & eD & ")"
    ws.Cells(1, 1).Font.Color = RGB(192, 0, 0)
Done_:
    Application.ScreenUpdating = True
End Sub

' Excel consumes a leading apostrophe as the text-prefix character even on
' text-formatted cells - double it so COBOL literals ('0001') display intact.
Private Sub PutText_(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal s As String)
    If Left$(s, 1) = "'" Then s = "'" & s
    ws.Cells(r, c).Value = s
End Sub

Private Function RenderCaseIo_(ByVal ws As Worksheet, ByVal cm As OrderedDict, ByVal startRow As Long) As Long
    Dim row As Long
    row = startRow

    ' case band: deep green with white text (same meaning-color as the
    ' normal-case bands on テストケース候補)
    ws.Cells(row, 1).Value = ChrW$(&H25A0) & " " & CStr(cm.Item("id")) & "（正常系シナリオ" & CLng(cm.Item("kindSerial")) & "）　　終了形態: " & CStr(cm.Item("termJp"))
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 5))
        .Interior.Color = RGB(55, 86, 35)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Cells(row, 1).Font.Bold = True
    row = row + 1

    ' column header: dark blue band with white text
    ws.Cells(row, 1).Value = "分類"
    ws.Cells(row, 2).Value = "項目／内容"
    ws.Cells(row, 3).Value = "設定値・想定値"
    ws.Cells(row, 4).Value = "実測値（記入）"
    ws.Cells(row, 5).Value = "判定（記入）"
    With ws.Range(ws.Cells(row, 1), ws.Cells(row, 5))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(46, 91, 143)
    End With
    ws.Range(ws.Cells(row, 4), ws.Cells(row, 5)).HorizontalAlignment = xlCenter
    row = row + 1
    Dim gridTop As Long
    gridTop = row

    Dim r As OrderedDict, v As Variant
    For Each r In cm.Item("lkIn")
        ws.Cells(row, 1).Value = "入力(1) LINKAGE"
        ws.Cells(row, 2).Value = CStr(r.Item("Item"))
        PutText_ ws, row, 3, CStr(r.Item("Note"))
        row = row + 1
    Next r
    For Each v In cm.Item("dbPre")
        ws.Cells(row, 1).Value = "入力(2) DB前提"
        PutText_ ws, row, 3, CStr(v)
        row = row + 1
    Next v
    For Each v In cm.Item("subPre")
        ws.Cells(row, 1).Value = "入力(3) サブ戻り値"
        PutText_ ws, row, 3, CStr(v)
        row = row + 1
    Next v
    ' input-class column band (light blue, navy text)
    If row > gridTop Then
        With ws.Range(ws.Cells(gridTop, 1), ws.Cells(row - 1, 1))
            .Interior.Color = RGB(220, 230, 241)
            .Font.Color = RGB(31, 78, 121)
            .Font.Bold = True
        End With
    End If
    If cm.Item("lkIn").Count = 0 And cm.Item("dbPre").Count = 0 And cm.Item("subPre").Count = 0 Then
        ws.Cells(row, 1).Value = "入力"
        ws.Cells(row, 3).Value = "（特記事項なし）"
        row = row + 1
    End If

    For Each r In cm.Item("outs")
        ws.Cells(row, 1).Value = "出力"
        With ws.Cells(row, 1)
            .Font.Color = RGB(31, 95, 165)
            .Font.Bold = True
        End With
        ws.Cells(row, 2).Value = CStr(r.Item("Item"))
        PutText_ ws, row, 3, CStr(r.Item("Val"))
        ws.Range(ws.Cells(row, 4), ws.Cells(row, 5)).Interior.Color = RGB(255, 253, 231)
        row = row + 1
    Next r
    ws.Cells(row, 1).Value = "出力"
    With ws.Cells(row, 1)
        .Font.Color = RGB(31, 95, 165)
        .Font.Bold = True
    End With
    ws.Cells(row, 2).Value = "終了形態"
    ws.Cells(row, 3).Value = CStr(cm.Item("termJp"))
    row = row + 1

    For Each r In cm.Item("refs")
        ws.Cells(row, 1).Value = "参考（内部）"
        ws.Cells(row, 1).Font.Color = RGB(120, 120, 120)
        ws.Cells(row, 2).Value = CStr(r.Item("Item"))
        PutText_ ws, row, 3, CStr(r.Item("Val"))
        ws.Cells(row, 2).Font.Color = RGB(120, 120, 120)
        ws.Cells(row, 3).Font.Color = RGB(120, 120, 120)
        row = row + 1
    Next r

    ' block grid (header band through the last row, one range call)
    With ws.Range(ws.Cells(gridTop - 1, 1), ws.Cells(row - 1, 5)).Borders
        .LineStyle = xlContinuous
        .Color = RGB(184, 188, 196)
        .Weight = xlThin
    End With

    RenderCaseIo_ = row
End Function
