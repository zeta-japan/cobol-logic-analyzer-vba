Attribute VB_Name = "CobolParser"
' CobolParser - COBOL logic analyzer engine.
' Phase 1 scope: line normalization with prefix auto-detection, program
' name extraction, and a minimal analyze entry that produces the JSON
' summary { programName, lines, prefixDetected, prefixStyle, prefixRatio }.
' Subsequent phases add AST, paths, coverage, and call graph.

Option Explicit

Public Const PARSER_VERSION As String = "0.1.0"

Public Function Convert_CollapseSpaces(ByVal txt As String) As String
    Static rx As Object
    If rx Is Nothing Then
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "\s+"
        rx.Global = True
    End If
    Convert_CollapseSpaces = Trim$(rx.Replace(txt, " "))
End Function

Public Function Convert_StripTrailingPeriod(ByVal txt As String) As String
    Dim t As String
    t = Trim$(txt)
    If Len(t) > 0 Then
        If Right$(t, 1) = "." Then t = Left$(t, Len(t) - 1)
    End If
    Convert_StripTrailingPeriod = Trim$(t)
End Function

Public Function Convert_NormalizeCondition(ByVal cond As String) As String
    Convert_NormalizeCondition = Convert_CollapseSpaces(Convert_StripTrailingPeriod(cond))
End Function

' Returns OrderedDict { Lines, PrefixDetected, PrefixStyle, PrefixRatio }
' where Lines is a Collection of OrderedDict { Number, Raw, Text }.
Public Function Get_NormalizedCobolLines(ByVal source As String, Optional ByVal forcePrefix As String = "") As OrderedDict
    Dim rawLines() As String
    rawLines = Split(Replace(source, Chr$(13), ""), Chr$(10))

    Dim rxPrefixNum As Object, rxPrefixSep As Object, rxStandardHead As Object, rxComment As Object
    Set rxPrefixNum = CreateObject("VBScript.RegExp")
    rxPrefixNum.Pattern = "^\d{6}$"
    Set rxPrefixSep = CreateObject("VBScript.RegExp")
    rxPrefixSep.Pattern = "[\s*/]"
    Set rxStandardHead = CreateObject("VBScript.RegExp")
    rxStandardHead.Pattern = "^\s*\d{0,6}\s"
    Set rxComment = CreateObject("VBScript.RegExp")
    rxComment.Pattern = "^[*/]"

    Dim nonBlank As Long, prefixHit As Long
    Dim i As Long, rl As String
    For i = LBound(rawLines) To UBound(rawLines)
        rl = rawLines(i)
        If Len(Trim$(rl)) > 0 Then
            nonBlank = nonBlank + 1
            If Len(rl) >= 7 Then
                If rxPrefixNum.Test(Left$(rl, 6)) And rxPrefixSep.Test(Mid$(rl, 7, 1)) Then
                    prefixHit = prefixHit + 1
                End If
            End If
        End If
    Next i

    Dim ratio As Double
    If nonBlank > 0 Then ratio = prefixHit / nonBlank

    Dim style As String, detected As Boolean
    Select Case LCase$(forcePrefix)
        Case "prefixed":  style = "prefixed":  detected = True
        Case "standard":  style = "standard":  detected = False
        Case "none":      style = "none":      detected = False
        Case Else
            If ratio >= 0.6 Then
                style = "prefixed":  detected = True
            Else
                style = "standard":  detected = False
            End If
    End Select

    Dim lines As Collection
    Set lines = New Collection

    Dim idx As Long, stripped As String, headStandard As String, txt As String
    Dim entry As OrderedDict
    For i = LBound(rawLines) To UBound(rawLines)
        rl = rawLines(i)
        idx = idx + 1
        If style = "prefixed" Then
            If Len(rl) >= 6 Then stripped = Mid$(rl, 7) Else stripped = rl
        ElseIf style = "none" Then
            stripped = rl
        Else
            If Len(rl) >= 8 Then headStandard = Left$(rl, 8) Else headStandard = rl
            If Len(rl) > 6 And rxStandardHead.Test(headStandard) Then
                stripped = Mid$(rl, 7)
            Else
                stripped = rl
            End If
        End If
        txt = Convert_CollapseSpaces(stripped)
        If txt <> "" And Not rxComment.Test(txt) Then
            Set entry = New OrderedDict
            entry.Add "Number", idx
            entry.Add "Raw", rl
            entry.Add "Text", txt
            lines.Add entry
        End If
    Next i

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "Lines", lines
    result.Add "PrefixDetected", detected
    result.Add "PrefixStyle", style
    result.Add "PrefixRatio", Round(ratio, 3)
    Set Get_NormalizedCobolLines = result
End Function

' Extracts PROGRAM-ID, or first SECTION name, or "(NO-NAME)" if neither found.
Public Function Get_ProgramName(ByVal lines As Collection) As String
    Dim rxPid As Object, rxSec As Object, entry As OrderedDict, m As Object
    Set rxPid = CreateObject("VBScript.RegExp")
    rxPid.Pattern = "^PROGRAM-ID\.\s*([A-Z0-9][A-Z0-9_-]*)"
    rxPid.IgnoreCase = False
    Set rxSec = CreateObject("VBScript.RegExp")
    rxSec.Pattern = "^([A-Z0-9][A-Z0-9_-]+)\s+SECTION$"
    rxSec.IgnoreCase = False

    For Each entry In lines
        Set m = rxPid.Execute(entry.Item("Text"))
        If m.Count > 0 Then
            Get_ProgramName = UCase$(m.Item(0).SubMatches(0))
            Exit Function
        End If
    Next entry
    For Each entry In lines
        Set m = rxSec.Execute(entry.Item("Text"))
        If m.Count > 0 Then
            Get_ProgramName = UCase$(m.Item(0).SubMatches(0))
            Exit Function
        End If
    Next entry
    Get_ProgramName = "(NO-NAME)"
End Function

' Phase 1 orchestrator. Returns { summary: { programName, lines, prefix... } }.
Public Function Analyze_Phase1(ByVal source As String) As OrderedDict
    Dim norm As OrderedDict
    Set norm = Get_NormalizedCobolLines(source, "")

    Dim lines As Collection
    Set lines = norm.Item("Lines")

    Dim summary As OrderedDict
    Set summary = New OrderedDict
    summary.Add "programName", Get_ProgramName(lines)
    summary.Add "lines", lines.Count
    summary.Add "prefixDetected", norm.Item("PrefixDetected")
    summary.Add "prefixStyle", norm.Item("PrefixStyle")
    summary.Add "prefixRatio", norm.Item("PrefixRatio")

    Dim result As OrderedDict
    Set result = New OrderedDict
    result.Add "summary", summary
    Set Analyze_Phase1 = result
End Function
