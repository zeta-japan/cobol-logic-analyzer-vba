Attribute VB_Name = "CobolCalls"
' CobolCalls - ver2.0 extraction for the call/usage diagram.
'   Get_ExternalCalls(lines) : external CALL targets + their USING arguments
'   Get_DataAccess(lines)    : files/records accessed (OPEN/READ/WRITE/...),
'                              record names mapped to their file via FD
' Pure extraction (no Japanese, no rendering). Reuses CobolParser helpers.
' Kept in a separate ASCII module so edits never touch CobolParser.bas, which
' holds CP932 Japanese literals.

Option Explicit

' Returns a Collection of OrderedDict { program, args(Collection of String), line }.
Public Function Get_ExternalCalls(ByVal lines As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "^CALL\s+('([^']+)'|([A-Z0-9][A-Za-z0-9_-]*))(\s+USING\s+(.+))?$"
    rx.IgnoreCase = False

    Dim line As OrderedDict, txt As String, m As Object, sm As Object
    For Each line In lines
        txt = CobolParser.Convert_StripTrailingPeriod(line.Item("Text"))
        Set m = rx.Execute(txt)
        If m.Count > 0 Then
            Set sm = m.Item(0).SubMatches
            Dim prog As String
            If Len(sm.Item(1)) > 0 Then prog = sm.Item(1) Else prog = sm.Item(2)

            Dim args As Collection
            Set args = New Collection
            If Len(sm.Item(4)) > 0 Then
                Dim parts() As String, i As Long, tok As String
                parts = Split(sm.Item(4), " ")
                For i = LBound(parts) To UBound(parts)
                    tok = Trim$(parts(i))
                    If tok <> "" And Not IsArgKeyword_(tok) Then args.Add tok
                Next i
            End If

            Dim e As OrderedDict
            Set e = New OrderedDict
            e.Add "program", UCase$(prog)
            e.Add "args", args
            e.Add "line", line.Item("Number")
            result.Add e
        End If
    Next line
    Set Get_ExternalCalls = result
End Function

' Returns a Collection of OrderedDict { name, modes(Collection of String), line }.
' Record names used by WRITE/REWRITE/DELETE are mapped to their file (via FD).
Public Function Get_DataAccess(ByVal lines As Collection) As Collection
    Dim recToFile As OrderedDict
    Set recToFile = BuildRecordToFileMap_(lines)

    ' name -> modes collection; first-seen order preserved via 'names'
    Dim names As Collection, modeMap As OrderedDict, lineMap As OrderedDict
    Set names = New Collection
    Set modeMap = New OrderedDict
    Set lineMap = New OrderedDict

    Dim rxOpen As Object, rxIO As Object
    Set rxOpen = CreateObject("VBScript.RegExp")
    rxOpen.Pattern = "^OPEN\s+(INPUT|OUTPUT|I-O|EXTEND)\s+(.+)$"
    rxOpen.IgnoreCase = False
    Set rxIO = CreateObject("VBScript.RegExp")
    rxIO.Pattern = "^(READ|WRITE|REWRITE|DELETE|START|CLOSE)\s+([A-Z0-9][A-Za-z0-9_-]*)"
    rxIO.IgnoreCase = False

    Dim line As OrderedDict, txt As String, m As Object
    For Each line In lines
        txt = CobolParser.Convert_StripTrailingPeriod(line.Item("Text"))

        Set m = rxOpen.Execute(txt)
        If m.Count > 0 Then
            Dim openMode As String, files() As String, k As Long
            openMode = "OPEN-" & m.Item(0).SubMatches(0)
            files = Split(m.Item(0).SubMatches(1), " ")
            For k = LBound(files) To UBound(files)
                If Trim$(files(k)) <> "" Then
                    AddAccess_ names, modeMap, lineMap, Trim$(files(k)), openMode, CLng(line.Item("Number"))
                End If
            Next k
        Else
            Set m = rxIO.Execute(txt)
            If m.Count > 0 Then
                Dim verb As String, operand As String
                verb = m.Item(0).SubMatches(0)
                operand = m.Item(0).SubMatches(1)
                If recToFile.Exists(operand) Then operand = CStr(recToFile.Item(operand))
                AddAccess_ names, modeMap, lineMap, operand, verb, CLng(line.Item("Number"))
            End If
        End If
    Next line

    Dim result As Collection
    Set result = New Collection
    Dim nm As Variant, e As OrderedDict
    For Each nm In names
        Set e = New OrderedDict
        e.Add "name", CStr(nm)
        e.Add "modes", modeMap.Item(CStr(nm))
        e.Add "line", lineMap.Item(CStr(nm))
        result.Add e
    Next nm
    Set Get_DataAccess = result
End Function

Private Function BuildRecordToFileMap_(ByVal lines As Collection) As OrderedDict
    Dim map As OrderedDict
    Set map = New OrderedDict
    Dim rxFD As Object, rx01 As Object
    Set rxFD = CreateObject("VBScript.RegExp"): rxFD.Pattern = "^FD\s+([A-Z0-9][A-Za-z0-9_-]*)"
    Set rx01 = CreateObject("VBScript.RegExp"): rx01.Pattern = "^01\s+([A-Z0-9][A-Za-z0-9_-]*)"
    rxFD.IgnoreCase = False
    rx01.IgnoreCase = False

    Dim line As OrderedDict, txt As String, m As Object, pendingFD As String
    pendingFD = ""
    For Each line In lines
        txt = CobolParser.Convert_StripTrailingPeriod(line.Item("Text"))
        Set m = rxFD.Execute(txt)
        If m.Count > 0 Then
            pendingFD = m.Item(0).SubMatches(0)
        ElseIf pendingFD <> "" Then
            Set m = rx01.Execute(txt)
            If m.Count > 0 Then
                If Not map.Exists(m.Item(0).SubMatches(0)) Then map.Add m.Item(0).SubMatches(0), pendingFD
                pendingFD = ""
            End If
        End If
    Next line
    Set BuildRecordToFileMap_ = map
End Function

Private Sub AddAccess_(ByVal names As Collection, ByVal modeMap As OrderedDict, _
                       ByVal lineMap As OrderedDict, ByVal nm As String, _
                       ByVal mode As String, ByVal ln As Long)
    Dim modes As Collection
    If modeMap.Exists(nm) Then
        Set modes = modeMap.Item(nm)
    Else
        Set modes = New Collection
        modeMap.Add nm, modes
        lineMap.Add nm, ln
        names.Add nm
    End If
    Dim v As Variant, found As Boolean
    found = False
    For Each v In modes
        If v = mode Then found = True: Exit For
    Next v
    If Not found Then modes.Add mode
End Sub

Private Function IsArgKeyword_(ByVal tok As String) As Boolean
    Select Case UCase$(tok)
        Case "BY", "REFERENCE", "CONTENT", "VALUE": IsArgKeyword_ = True
        Case Else: IsArgKeyword_ = False
    End Select
End Function
