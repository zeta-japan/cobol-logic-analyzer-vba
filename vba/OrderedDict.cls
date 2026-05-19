VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OrderedDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OrderedDict - ordered key/value container.
' Preserves insertion order of keys. Used by JsonWriter to guarantee field
' order in the emitted JSON output (Scripting.Dictionary does not guarantee
' key order).

Option Explicit

Private mKeys As Collection
Private mItems As Object

Private Sub Class_Initialize()
    Set mKeys = New Collection
    Set mItems = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Add(ByVal key As String, ByVal value As Variant)
    If mItems.Exists(key) Then
        If IsObject(value) Then
            Set mItems(key) = value
        Else
            mItems(key) = value
        End If
    Else
        mKeys.Add key
        If IsObject(value) Then
            mItems.Add key, value
        Else
            mItems.Add key, value
        End If
    End If
End Sub

Public Property Get Item(ByVal key As String) As Variant
    If mItems.Exists(key) Then
        If IsObject(mItems(key)) Then
            Set Item = mItems(key)
        Else
            Item = mItems(key)
        End If
    End If
End Property

Public Function Exists(ByVal key As String) As Boolean
    Exists = mItems.Exists(key)
End Function

Public Function Keys() As Collection
    Set Keys = mKeys
End Function

Public Property Get Count() As Long
    Count = mItems.Count
End Property
