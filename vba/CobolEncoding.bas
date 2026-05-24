Attribute VB_Name = "CobolEncoding"
' CobolEncoding - read COBOL source text with explicit or auto-detected charset.
' Wraps ADODB.Stream. Returns the file content as a VBA String (UTF-16 LE
' internally; the original bytes are decoded according to charset).

Option Explicit

' Named ReadCobolSource (not ReadAllText) to avoid colliding with
' JsonParser.ReadAllText, which CobolLogicViewer calls unqualified.
Public Function ReadCobolSource(ByVal path As String, Optional ByVal charset As String = "auto") As String
    If Len(Dir(path)) = 0 Then
        Err.Raise 53, "CobolEncoding.ReadCobolSource", "File not found: " & path
    End If
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' adTypeText
    st.charset = ResolveCharset(path, charset)
    st.Open
    st.LoadFromFile path
    ReadCobolSource = st.ReadText
    st.Close
End Function

' Write a UTF-8 (or other) text file. Used by Main.AnalyzeAndBuild to drop the
' JSON file that CobolLogicViewer.BuildCobolReport reads back.
Public Sub WriteAllText(ByVal path As String, ByVal text As String, Optional ByVal charset As String = "utf-8")
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' adTypeText
    Select Case LCase$(charset)
        Case "utf-8", "utf8":   st.charset = "utf-8"
        Case "cp932", "shift_jis", "shift-jis", "sjis": st.charset = "shift_jis"
        Case "utf-16", "utf-16le", "unicode": st.charset = "unicode"
        Case Else: st.charset = charset
    End Select
    st.Open
    st.WriteText text
    st.SaveToFile path, 2 ' adSaveCreateOverWrite
    st.Close
End Sub

Private Function ResolveCharset(ByVal path As String, ByVal req As String) As String
    Select Case LCase$(req)
        Case "utf-8", "utf8":                       ResolveCharset = "utf-8"
        Case "cp932", "shift_jis", "shift-jis", "sjis": ResolveCharset = "shift_jis"
        Case "utf-16", "utf-16le", "unicode":       ResolveCharset = "unicode"
        Case "auto", "":                            ResolveCharset = DetectCharset(path)
        Case Else:                                  ResolveCharset = req
    End Select
End Function

' Detect charset from the raw bytes. BOM wins; otherwise validate the byte
' stream as UTF-8 (COBOL sources here are UTF-8 without BOM). Only if that
' fails do we fall back to shift_jis (CP932), since misreading UTF-8 Japanese
' as CP932 lets a comment's trailing newline get swallowed and merges lines.
Private Function DetectCharset(ByVal path As String) As String
    Dim bin As Object, bytes() As Byte, raw As Variant
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1 ' adTypeBinary
    bin.Open
    bin.LoadFromFile path
    raw = bin.Read ' adReadAll
    bin.Close

    Dim n As Long
    n = -1
    On Error Resume Next
    bytes = raw
    n = UBound(bytes) + 1
    On Error GoTo 0
    If n <= 0 Then
        DetectCharset = "shift_jis"
        Exit Function
    End If

    If n >= 3 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            DetectCharset = "utf-8"
            Exit Function
        End If
    End If
    If n >= 2 Then
        If bytes(0) = &HFF And bytes(1) = &HFE Then
            DetectCharset = "unicode"
            Exit Function
        End If
    End If

    If IsValidUtf8(bytes, n) Then
        DetectCharset = "utf-8"
    Else
        DetectCharset = "shift_jis"
    End If
End Function

' True if the bytes are a well-formed UTF-8 stream. Pure ASCII counts as valid
' (it decodes identically either way). A genuine shift_jis stream with Japanese
' fails here because its trail bytes (0x40-0x7E) are not valid UTF-8 continuation
' bytes (which must be 0x80-0xBF).
Private Function IsValidUtf8(ByRef b() As Byte, ByVal n As Long) As Boolean
    Dim i As Long, c As Long, trail As Long, k As Long
    i = 0
    Do While i < n
        c = b(i)
        If c <= &H7F Then
            trail = 0
        ElseIf c >= &HC2 And c <= &HDF Then
            trail = 1
        ElseIf c >= &HE0 And c <= &HEF Then
            trail = 2
        ElseIf c >= &HF0 And c <= &HF4 Then
            trail = 3
        Else
            IsValidUtf8 = False
            Exit Function
        End If
        For k = 1 To trail
            If i + k >= n Then
                IsValidUtf8 = False
                Exit Function
            End If
            If b(i + k) < &H80 Or b(i + k) > &HBF Then
                IsValidUtf8 = False
                Exit Function
            End If
        Next k
        i = i + trail + 1
    Loop
    IsValidUtf8 = True
End Function
