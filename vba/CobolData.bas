Attribute VB_Name = "CobolData"
' CobolData - ver2.0 feature (2): extract data items / arguments from the
' DATA DIVISION (WORKING-STORAGE / LINKAGE / FILE / LOCAL-STORAGE).
'   Get_DataItems(lines) : Collection of OrderedDict
'       { section, level, name, pic, picType, picLen, occurs, value,
'         redefines, isCopy, line }
' PIC is classified to a type (num / signed-num / decimal / alnum / group)
' and a display length. COPY lines are recorded as isCopy=True (the copybook
' body is external, so its items are not expanded).
' ASCII only (no Japanese) - kept separate from the CP932 modules.

Option Explicit

Public Function Get_DataItems(ByVal lines As Collection) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim rxSec As Object, rxItem As Object, rxPic As Object, rxOcc As Object
    Dim rxVal As Object, rxRed As Object, rxCopy As Object
    Set rxSec = MakeRx_("^(WORKING-STORAGE|LINKAGE|FILE|LOCAL-STORAGE)\s+SECTION$")
    Set rxItem = MakeRx_("^(\d{2})\s+([A-Z0-9][A-Za-z0-9_-]*)")
    Set rxPic = MakeRx_("\bPIC(TURE)?\s+(\S+)")
    Set rxOcc = MakeRx_("\bOCCURS\s+(\d+)")
    Set rxVal = MakeRx_("\bVALUE\s+(.+)$")
    Set rxRed = MakeRx_("\bREDEFINES\s+([A-Za-z0-9_-]+)")
    Set rxCopy = MakeRx_("^COPY\s+([A-Z0-9][A-Za-z0-9_-]*)")

    Dim section As String
    section = ""
    Dim line As OrderedDict, txt As String, m As Object
    For Each line In lines
        txt = CobolParser.Convert_StripTrailingPeriod(line.Item("Text"))

        Set m = rxSec.Execute(txt)
        If m.Count > 0 Then
            section = m.Item(0).SubMatches(0)
            GoTo NextItem
        End If

        Set m = rxCopy.Execute(txt)
        If m.Count > 0 Then
            result.Add MakeItem_(section, "", m.Item(0).SubMatches(0), "", "copy", 0, "", "", "", True, CLng(line.Item("Number")))
            GoTo NextItem
        End If

        Set m = rxItem.Execute(txt)
        If m.Count > 0 Then
            Dim lvl As String, nm As String
            lvl = m.Item(0).SubMatches(0)
            nm = m.Item(0).SubMatches(1)
            If UCase$(nm) = "SECTION" Or UCase$(nm) = "DIVISION" Then GoTo NextItem

            Dim pic As String, occ As String, val As String, red As String
            pic = "": occ = "": val = "": red = ""
            Set m = rxPic.Execute(txt): If m.Count > 0 Then pic = m.Item(0).SubMatches(1)
            Set m = rxOcc.Execute(txt): If m.Count > 0 Then occ = m.Item(0).SubMatches(0)
            Set m = rxVal.Execute(txt): If m.Count > 0 Then val = Trim$(m.Item(0).SubMatches(0))
            Set m = rxRed.Execute(txt): If m.Count > 0 Then red = m.Item(0).SubMatches(0)

            Dim pType As String, pLen As Long
            If pic = "" Then
                pType = "group": pLen = 0
            Else
                pType = PicType_(pic): pLen = PicLen_(pic)
            End If

            result.Add MakeItem_(section, lvl, nm, pic, pType, pLen, occ, val, red, False, CLng(line.Item("Number")))
        End If
NextItem:
    Next line
    Set Get_DataItems = result
End Function

' Classify a PIC string to a coarse type.
Public Function PicType_(ByVal pic As String) As String
    Dim p As String
    p = UCase$(StripDot_(pic))
    If InStr(p, "S") > 0 Then
        PicType_ = "signed-num"
    ElseIf InStr(p, "V") > 0 Then
        PicType_ = "decimal"
    ElseIf InStr(p, "9") > 0 Then
        PicType_ = "num"
    ElseIf InStr(p, "X") > 0 Or InStr(p, "A") > 0 Then
        PicType_ = "alnum"
    Else
        PicType_ = "other"
    End If
End Function

' Display length: sum of position counts (X/A/9/Z/P), honoring "(n)". S and V
' do not occupy a display position.
Public Function PicLen_(ByVal pic As String) As Long
    Dim p As String, i As Long, ch As String, total As Long, closePos As Long, numStr As String
    p = UCase$(StripDot_(pic))
    total = 0: i = 1
    Do While i <= Len(p)
        ch = Mid$(p, i, 1)
        If InStr("XA9ZP", ch) > 0 Or ch = "S" Or ch = "V" Then
            If i < Len(p) And Mid$(p, i + 1, 1) = "(" Then
                closePos = InStr(i, p, ")")
                If closePos > 0 Then
                    numStr = Mid$(p, i + 2, closePos - i - 2)
                    If ch <> "S" And ch <> "V" Then total = total + CLng(Val(numStr))
                    i = closePos + 1
                Else
                    i = i + 1
                End If
            Else
                If ch <> "S" And ch <> "V" And ch <> "P" Then total = total + 1
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
    PicLen_ = total
End Function

' A short test-value hint for the item (used by the items sheet / future driver).
Public Function ValueHint_(ByVal picType As String, ByVal picLen As Long) As String
    Select Case picType
        Case "num", "signed-num", "decimal"
            ValueHint_ = "0 / " & String$(IIf(picLen > 0 And picLen < 18, picLen, 1), "9")
        Case "alnum"
            ValueHint_ = "SPACE / '" & String$(IIf(picLen > 0 And picLen < 10, picLen, 1), "X") & "'"
        Case Else
            ValueHint_ = ""
    End Select
End Function

Private Function StripDot_(ByVal s As String) As String
    If Right$(s, 1) = "." Then
        StripDot_ = Left$(s, Len(s) - 1)
    Else
        StripDot_ = s
    End If
End Function

Private Function MakeRx_(ByVal pattern As String) As Object
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = pattern
    r.IgnoreCase = False
    r.Global = False
    Set MakeRx_ = r
End Function

Private Function MakeItem_(ByVal section As String, ByVal level As String, ByVal name As String, _
                           ByVal pic As String, ByVal picType As String, ByVal picLen As Long, _
                           ByVal occurs As String, ByVal value As String, ByVal redefines As String, _
                           ByVal isCopy As Boolean, ByVal lineNo As Long) As OrderedDict
    Dim e As OrderedDict
    Set e = New OrderedDict
    e.Add "section", section
    e.Add "level", level
    e.Add "name", name
    e.Add "pic", pic
    e.Add "picType", picType
    e.Add "picLen", picLen
    e.Add "occurs", occurs
    e.Add "value", value
    e.Add "redefines", redefines
    e.Add "isCopy", isCopy
    e.Add "line", lineNo
    Set MakeItem_ = e
End Function
