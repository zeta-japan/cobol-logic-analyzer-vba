Attribute VB_Name = "CobolEncoding"
' CobolEncoding - read COBOL source text with explicit or auto-detected charset.
' Wraps ADODB.Stream. Returns the file content as a VBA String (UTF-16 LE
' internally; the original bytes are decoded according to charset).

Option Explicit

Public Function ReadAllText(ByVal path As String, Optional ByVal charset As String = "auto") As String
    If Len(Dir(path)) = 0 Then
        Err.Raise 53, "CobolEncoding.ReadAllText", "File not found: " & path
    End If
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2 ' adTypeText
    st.charset = ResolveCharset(path, charset)
    st.Open
    st.LoadFromFile path
    ReadAllText = st.ReadText
    st.Close
End Function

Private Function ResolveCharset(ByVal path As String, ByVal req As String) As String
    Select Case LCase$(req)
        Case "utf-8", "utf8":                       ResolveCharset = "utf-8"
        Case "cp932", "shift_jis", "shift-jis", "sjis": ResolveCharset = "shift_jis"
        Case "utf-16", "utf-16le", "unicode":       ResolveCharset = "unicode"
        Case "auto", "":                            ResolveCharset = DetectCharsetByBom(path)
        Case Else:                                  ResolveCharset = req
    End Select
End Function

Private Function DetectCharsetByBom(ByVal path As String) As String
    Dim bin As Object, head As Variant
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1 ' adTypeBinary
    bin.Open
    bin.LoadFromFile path
    On Error Resume Next
    head = bin.Read(4)
    On Error GoTo 0
    bin.Close

    Dim b0 As Long, b1 As Long, b2 As Long, n As Long
    On Error Resume Next
    n = UBound(head) + 1
    On Error GoTo 0
    If n >= 3 Then
        b0 = head(0): b1 = head(1): b2 = head(2)
        If b0 = &HEF And b1 = &HBB And b2 = &HBF Then
            DetectCharsetByBom = "utf-8"
            Exit Function
        End If
    End If
    If n >= 2 Then
        b0 = head(0): b1 = head(1)
        If b0 = &HFF And b1 = &HFE Then
            DetectCharsetByBom = "unicode"
            Exit Function
        End If
    End If
    DetectCharsetByBom = "shift_jis"
End Function
