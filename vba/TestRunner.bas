Attribute VB_Name = "TestRunner"
' TestRunner - lightweight in-workbook test framework.
' Provides Assert_True / Assert_Equal helpers and renders results to a
' TestResults sheet plus the Immediate window. Run_All_Tests is the entry.

Option Explicit

Private mPass As Long
Private mFail As Long
Private mLog As Collection

Public Sub Test_Begin()
    mPass = 0
    mFail = 0
    Set mLog = New Collection
End Sub

Public Sub Assert_True(ByVal cond As Boolean, ByVal msg As String)
    If cond Then
        mPass = mPass + 1
        Log_ "[PASS] " & msg
    Else
        mFail = mFail + 1
        Log_ "[FAIL] " & msg
    End If
End Sub

Public Sub Assert_Equal(ByVal expected As Variant, ByVal actual As Variant, ByVal msg As String)
    Dim equal As Boolean
    On Error Resume Next
    equal = (expected = actual)
    On Error GoTo 0
    If equal Then
        mPass = mPass + 1
        Log_ "[PASS] " & msg
    Else
        mFail = mFail + 1
        Log_ "[FAIL] " & msg & "  expected=" & SafeToString_(expected) & "  actual=" & SafeToString_(actual)
    End If
End Sub

Public Sub Test_End()
    Log_ "---- " & mPass & " passed, " & mFail & " failed"
    Render_Log
End Sub

Public Sub Run_All_Tests()
    Test_Begin
    Log_ "Workbook path: " & ThisWorkbook.path
    Test_Phase1.Run_All
    Test_End
End Sub

' Run one named test sub via Application.Run, catching any unhandled error
' and converting it to a single [FAIL] line. Lets the rest of the suite run.
Public Sub Run_One(ByVal testName As String)
    On Error Resume Next
    Application.Run testName
    If Err.Number <> 0 Then
        Assert_True False, testName & " threw error #" & Err.Number & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub Log_(ByVal line As String)
    Debug.Print line
    If mLog Is Nothing Then Set mLog = New Collection
    mLog.Add line
End Sub

Private Function SafeToString_(ByVal v As Variant) As String
    On Error Resume Next
    If IsObject(v) Then
        SafeToString_ = "<" & TypeName(v) & ">"
    ElseIf IsNull(v) Then
        SafeToString_ = "<Null>"
    ElseIf IsEmpty(v) Then
        SafeToString_ = "<Empty>"
    Else
        SafeToString_ = CStr(v)
    End If
    On Error GoTo 0
End Function

Private Sub Render_Log()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TestResults")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "TestResults"
    End If
    ws.Cells.Clear
    Dim r As Long, v As Variant
    r = 1
    For Each v In mLog
        ws.Cells(r, 1).value = v
        r = r + 1
    Next v
End Sub
