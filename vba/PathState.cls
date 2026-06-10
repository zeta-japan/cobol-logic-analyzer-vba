VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' PathState - lightweight accumulator for one (partial) execution path during
' Expand_NodeSequence. The fields are typed As Object because they carry
' ConsList heads (immutable shared lists, O(1) extension) DURING enumeration
' and are swapped for ordered Collections when the state is materialized at
' the end - consumers (test-case builder, tests) always see Collections.

Option Explicit

Public Conditions As Object
Public Actions As Object
Public Lines As Object
Public BranchIds As Object
