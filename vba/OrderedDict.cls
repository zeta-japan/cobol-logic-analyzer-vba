VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderedDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OrderedDict - insertion-ordered key/value container backed by two plain
' VBA arrays (parallel mKeys/mVals).
'
' Rationale: on locked-down machines both alternatives are pathologically slow.
'   - Scripting.Dictionary: CreateObject costs ~60ms per instance (AV scans the
'     scrrun.dll COM instantiation), and we build hundreds of these.
'   - intrinsic VBA Collection: keyed Item + error-trap Exists are very slow.
' Arrays have neither problem: "New OrderedDict" allocates nothing external, and
' lookups are a short linear scan. Our dicts hold at most ~15 keys, so O(n) per
' op is trivially fast and there is no COM / no error trapping anywhere.

Option Explicit

Private mKeys() As String
Private mVals() As Variant
Private mCount As Long

Private Sub Class_Initialize()
    mCount = 0
    ReDim mKeys(0 To 7)
    ReDim mVals(0 To 7)
End Sub

Private Function IndexOf(ByVal key As String) As Long
    Dim i As Long
    For i = 0 To mCount - 1
        If mKeys(i) = key Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = -1
End Function

Public Sub Add(ByVal key As String, ByVal value As Variant)
    Dim idx As Long
    idx = IndexOf(key)
    If idx >= 0 Then
        If IsObject(value) Then Set mVals(idx) = value Else mVals(idx) = value
        Exit Sub
    End If
    If mCount > UBound(mKeys) Then
        ReDim Preserve mKeys(0 To (mCount * 2) - 1)
        ReDim Preserve mVals(0 To (mCount * 2) - 1)
    End If
    mKeys(mCount) = key
    If IsObject(value) Then Set mVals(mCount) = value Else mVals(mCount) = value
    mCount = mCount + 1
End Sub

Public Property Get Item(ByVal key As String) As Variant
    Dim idx As Long
    idx = IndexOf(key)
    If idx >= 0 Then
        If IsObject(mVals(idx)) Then Set Item = mVals(idx) Else Item = mVals(idx)
    End If
End Property

Public Function Exists(ByVal key As String) As Boolean
    Exists = (IndexOf(key) >= 0)
End Function

' Keys in insertion order, as a 1-based Collection (callers use For Each and
' .Item(n)).
Public Function Keys() As Collection
    Dim c As Collection, i As Long
    Set c = New Collection
    For i = 0 To mCount - 1
        c.Add mKeys(i)
    Next i
    Set Keys = c
End Function

Public Property Get Count() As Long
    Count = mCount
End Property
