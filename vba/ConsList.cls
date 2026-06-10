VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ConsList - immutable singly-linked list node ("cons cell") for the path
' enumerators. Appending is O(1) (new head pointing at the old one) and
' "cloning" a list is O(1) (share the head), which turns the per-fork /
' per-action full-Collection copies in CobolFlow and Expand_NodeSequence
' from quadratic into linear work. Lists are only materialized back into
' Collections for the handful of finally-selected paths.
' The head node is the LAST element; walk Prev to reach earlier elements.

Option Explicit

Public V As Variant      ' the element (value or object reference)
Public Prev As ConsList  ' earlier elements; Nothing = list start
Public N As Long         ' length of the list ending at this node
