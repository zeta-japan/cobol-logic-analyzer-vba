VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderedDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OrderedDict - insertion-ordered key/value container.
'
' Backed by two intrinsic VBA Collections (no Scripting.Dictionary). This is
' deliberate: CreateObject("Scripting.Dictionary") can cost tens of
' milliseconds per instance on locked-down machines (AV scans each COM
' instantiation), and the analyzer builds hundreds of these. "New Collection"
' is in-process and effectively free, so creation is ~instant.
'
' mKeys keeps the keys in insertion order (for stable JSON output); mVals is a
' keyed Collection for O(1)-ish lookup. Keys are case-insensitive (Collection
' semantics) and must be unique - fine for our fixed JSON field names.

Option Explicit

Private mKeys As Collection
Private mVals As Collection

Private Sub Class_Initialize()
    Set mKeys = New Collection
    Set mVals = New Collection
End Sub

Public Sub Add(ByVal key As String, ByVal value As Variant)
    If Exists(key) Then
        mVals.Remove key            ' overwrite existing value, keep key order
    Else
        mKeys.Add key
    End If
    If IsObject(value) Then
        mVals.Add value, key
    Else
        mVals.Add value, key
    End If
End Sub

Public Property Get Item(ByVal key As String) As Variant
    On Error Resume Next
    If IsObject(mVals.Item(key)) Then
        Set Item = mVals.Item(key)
    Else
        Item = mVals.Item(key)
    End If
    On Error GoTo 0
End Property

Public Function Exists(ByVal key As String) As Boolean
    Dim probe As Boolean
    On Error Resume Next
    probe = IsObject(mVals.Item(key))   ' raises if key absent
    Exists = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function Keys() As Collection
    Set Keys = mKeys
End Function

Public Property Get Count() As Long
    Count = mKeys.Count
End Property
