Attribute VB_Name = "JsonParser"
' =========================================================================
' JsonParser - JSON parser and file/sheet helpers.
'
' Public API:
'   ReadAllText(path)  : Read a text file as UTF-8.
'   EnsureSheet(name)  : Get or create a worksheet by name.
'   ParseJson(text)    : Parse JSON text into Scripting.Dictionary / Collection.
'
' Dependencies:
'   - Reference: Microsoft Scripting Runtime
'   - ADODB.Stream (Windows built-in)
' =========================================================================
Option Explicit

Private json_pos As Long
Private json_src As String

' --- ファイル / シートヘルパー -------------------------------------------

Public Function ReadAllText(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "UTF-8"
    st.Open
    st.LoadFromFile path
    ReadAllText = st.ReadText(-1) ' adReadAll
    st.Close
End Function

Public Function EnsureSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = name
    End If
    Set EnsureSheet = ws
End Function

' --- JSON パーサー --------------------------------------------------------
' オブジェクト -> Scripting.Dictionary
' 配列         -> System.Collections.ArrayList
' 文字列       -> String   (\uXXXX 含む全エスケープ対応)
' 数値         -> Double
' true/false   -> Boolean
' null         -> Null

Public Function ParseJson(ByVal s As String) As Object
    json_src = s
    json_pos = 1
    SkipWs
    Set ParseJson = ParseValue
End Function

Private Function ParseValue() As Variant
    SkipWs
    Dim ch As String
    ch = Mid$(json_src, json_pos, 1)
    Select Case ch
        Case "{":  Set ParseValue = ParseObject
        Case "[":  Set ParseValue = ParseArray
        Case """": ParseValue = ParseString
        Case "t", "f": ParseValue = ParseBool
        Case "n":  ParseValue = ParseNull
        Case Else: ParseValue = ParseNumber
    End Select
End Function

Private Function ParseObject() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    json_pos = json_pos + 1 ' {
    SkipWs
    If Mid$(json_src, json_pos, 1) = "}" Then
        json_pos = json_pos + 1
        Set ParseObject = d
        Exit Function
    End If
    Do
        SkipWs
        Dim key As String
        key = ParseString
        SkipWs
        json_pos = json_pos + 1 ' :
        Dim val As Variant
        If IsObjectStart Then
            Set val = ParseValue
            d.Add key, val
        Else
            val = ParseValue
            d.Add key, val
        End If
        SkipWs
        If Mid$(json_src, json_pos, 1) = "," Then
            json_pos = json_pos + 1
        Else
            Exit Do
        End If
    Loop
    json_pos = json_pos + 1 ' }
    Set ParseObject = d
End Function

Private Function ParseArray() As Object
    Dim col As Collection
    Set col = New Collection
    json_pos = json_pos + 1 ' [
    SkipWs
    If Mid$(json_src, json_pos, 1) = "]" Then
        json_pos = json_pos + 1
        Set ParseArray = col
        Exit Function
    End If
    Do
        SkipWs
        Dim val As Variant
        If IsObjectStart Then
            Set val = ParseValue
            col.Add val
        Else
            val = ParseValue
            col.Add val
        End If
        SkipWs
        If Mid$(json_src, json_pos, 1) = "," Then
            json_pos = json_pos + 1
        Else
            Exit Do
        End If
    Loop
    json_pos = json_pos + 1 ' ]
    Set ParseArray = col
End Function

Private Function IsObjectStart() As Boolean
    SkipWs
    Dim ch As String
    ch = Mid$(json_src, json_pos, 1)
    IsObjectStart = (ch = "{" Or ch = "[")
End Function

Private Function ParseString() As String
    json_pos = json_pos + 1 ' opening "
    Dim sb As String
    sb = ""
    Do While json_pos <= Len(json_src)
        Dim ch As String
        ch = Mid$(json_src, json_pos, 1)
        If ch = "\" Then
            Dim esc As String
            esc = Mid$(json_src, json_pos + 1, 1)
            Select Case esc
                Case """": sb = sb & """":     json_pos = json_pos + 2
                Case "\":  sb = sb & "\":      json_pos = json_pos + 2
                Case "/":  sb = sb & "/":      json_pos = json_pos + 2
                Case "b":  sb = sb & Chr$(8):  json_pos = json_pos + 2
                Case "f":  sb = sb & Chr$(12): json_pos = json_pos + 2
                Case "n":  sb = sb & vbLf:     json_pos = json_pos + 2
                Case "r":  sb = sb & vbCr:     json_pos = json_pos + 2
                Case "t":  sb = sb & vbTab:    json_pos = json_pos + 2
                Case "u":
                    ' "&H0" を前置して 5 桁にし、16bit 符号付き解釈を回避
                    ' (漢字など 0x8000 以上のコードポイントを正しく扱う)
                    Dim hex4 As String
                    hex4 = Mid$(json_src, json_pos + 2, 4)
                    sb = sb & ChrW$(CLng("&H0" & hex4))
                    json_pos = json_pos + 6
                Case Else: sb = sb & esc:      json_pos = json_pos + 2
            End Select
        ElseIf ch = """" Then
            json_pos = json_pos + 1
            ParseString = sb
            Exit Function
        Else
            sb = sb & ch
            json_pos = json_pos + 1
        End If
    Loop
    ParseString = sb
End Function

Private Function ParseNumber() As Double
    Dim startPos As Long
    startPos = json_pos
    Do While json_pos <= Len(json_src) And InStr("0123456789-+.eE", Mid$(json_src, json_pos, 1)) > 0
        json_pos = json_pos + 1
    Loop
    ParseNumber = CDbl(Mid$(json_src, startPos, json_pos - startPos))
End Function

Private Function ParseBool() As Boolean
    If Mid$(json_src, json_pos, 4) = "true" Then
        ParseBool = True:  json_pos = json_pos + 4
    Else
        ParseBool = False: json_pos = json_pos + 5
    End If
End Function

Private Function ParseNull() As Variant
    json_pos = json_pos + 4
    ParseNull = Null
End Function

Private Sub SkipWs()
    Do While json_pos <= Len(json_src) And InStr(" " & vbTab & vbCr & vbLf, Mid$(json_src, json_pos, 1)) > 0
        json_pos = json_pos + 1
    Loop
End Sub
