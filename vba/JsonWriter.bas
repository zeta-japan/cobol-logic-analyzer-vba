Attribute VB_Name = "JsonWriter"
' JsonWriter - serialize VBA values to JSON text.
' Handles: OrderedDict (insertion order), Scripting.Dictionary, Collection,
'          VBA arrays, String, Long/Double, Boolean, Null/Empty/Nothing.
' Strings: \" \\ \b \t \n \f \r and control chars escape; non-ASCII passed
' through verbatim (matches PS ConvertTo-Json default behavior, with
' downstream encoding handled by the caller writing UTF-8).

Option Explicit

Public Function WriteJson(ByVal value As Variant) As String
    WriteJson = WriteValue(value)
End Function

Private Function WriteValue(ByVal value As Variant) As String
    If IsObject(value) Then
        If value Is Nothing Then
            WriteValue = "null"
            Exit Function
        End If
        If TypeOf value Is OrderedDict Then
            WriteValue = WriteOrderedDict(value)
            Exit Function
        End If
        Select Case TypeName(value)
            Case "Dictionary"
                WriteValue = WriteDictionary(value)
            Case "Collection"
                WriteValue = WriteCollection(value)
            Case Else
                WriteValue = "null"
        End Select
        Exit Function
    End If

    If IsArray(value) Then
        WriteValue = WriteArray(value)
        Exit Function
    End If

    If IsNull(value) Or IsEmpty(value) Then
        WriteValue = "null"
        Exit Function
    End If

    Select Case VarType(value)
        Case vbString
            WriteValue = WriteString(CStr(value))
        Case vbBoolean
            If CBool(value) Then WriteValue = "true" Else WriteValue = "false"
        Case vbInteger, vbLong, vbByte
            WriteValue = CStr(CLng(value))
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            WriteValue = FormatDouble_(CDbl(value))
        Case Else
            WriteValue = WriteString(CStr(value))
    End Select
End Function

Private Function WriteOrderedDict(ByVal od As OrderedDict) As String
    Dim parts As String, key As Variant, first As Boolean
    parts = "{"
    first = True
    For Each key In od.Keys
        If Not first Then parts = parts & ","
        parts = parts & WriteString(CStr(key)) & ":" & WriteValue(od.Item(CStr(key)))
        first = False
    Next key
    WriteOrderedDict = parts & "}"
End Function

Private Function WriteDictionary(ByVal d As Object) As String
    Dim parts As String, k As Variant, first As Boolean
    parts = "{"
    first = True
    For Each k In d.Keys
        If Not first Then parts = parts & ","
        parts = parts & WriteString(CStr(k)) & ":" & WriteValue(d(k))
        first = False
    Next k
    WriteDictionary = parts & "}"
End Function

Private Function WriteCollection(ByVal c As Collection) As String
    Dim parts As String, i As Long, first As Boolean
    parts = "["
    first = True
    For i = 1 To c.Count
        If Not first Then parts = parts & ","
        parts = parts & WriteValue(c.Item(i))
        first = False
    Next i
    WriteCollection = parts & "]"
End Function

Private Function WriteArray(ByVal arr As Variant) As String
    Dim parts As String, i As Long, first As Boolean
    Dim lo As Long, hi As Long
    parts = "["
    first = True
    On Error Resume Next
    lo = LBound(arr)
    hi = UBound(arr)
    On Error GoTo 0
    If hi >= lo Then
        For i = lo To hi
            If Not first Then parts = parts & ","
            parts = parts & WriteValue(arr(i))
            first = False
        Next i
    End If
    WriteArray = parts & "]"
End Function

Private Function WriteString(ByVal s As String) As String
    Dim sb As String, i As Long, ch As String, code As Long
    sb = """"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34: sb = sb & "\"""
            Case 92: sb = sb & "\\"
            Case 8:  sb = sb & "\b"
            Case 9:  sb = sb & "\t"
            Case 10: sb = sb & "\n"
            Case 12: sb = sb & "\f"
            Case 13: sb = sb & "\r"
            Case Is < 32: sb = sb & "\u" & PadHex_(code, 4)
            Case Else: sb = sb & ch
        End Select
    Next i
    WriteString = sb & """"
End Function

Private Function FormatDouble_(ByVal v As Double) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, ",", ".")
    FormatDouble_ = s
End Function

Private Function PadHex_(ByVal n As Long, ByVal width As Long) As String
    Dim s As String
    s = Hex$(n)
    Do While Len(s) < width
        s = "0" & s
    Loop
    PadHex_ = LCase$(s)
End Function
