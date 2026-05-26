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
' Expand_NodeSequence. Uses plain Collection fields with direct access instead
' of OrderedDict, because the enumeration creates thousands of these and
' OrderedDict's per-instance CreateObject("Scripting.Dictionary") plus the COM
' round-trips on every .Item access dominate the runtime.

Option Explicit

Public Conditions As Collection
Public Actions As Collection
Public Lines As Collection
Public BranchIds As Collection
