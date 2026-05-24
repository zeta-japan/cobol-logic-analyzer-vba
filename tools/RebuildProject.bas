' RebuildProject - one-shot helper to reload all VBA modules from the vba/ folder.
'
' Why: VBE's "File > Import" does NOT overwrite an existing module; it creates
' a duplicate (CobolParser1, TestRunner1, ...). That leaves stale copies and
' name clashes. This macro removes every standard/class module and re-imports
' the canonical set from disk, so the workbook always matches the repo.
'
' HOW TO USE (one time):
'   1. In Excel: File > Options > Trust Center > Trust Center Settings >
'      Macro Settings > check "Trust access to the VBA project object model".
'   2. In VBE, open the "ThisWorkbook" object (under Microsoft Excel Objects)
'      and PASTE the RebuildVbaProject sub below into it.
'      (It must live in ThisWorkbook so it does not delete itself.)
'   3. Run RebuildVbaProject (F5 while the cursor is inside it).
'   4. Delete the pasted sub afterwards if you like, then run Run_All_Tests.
'
' NOTE: this is a DEV convenience only. The shipped .xlsm for the client needs
' none of this; it is only for keeping the dev workbook in sync with git.

Public Sub RebuildVbaProject()
    Dim proj As Object
    On Error GoTo NoTrust
    Set proj = ThisWorkbook.VBProject
    On Error GoTo 0

    Dim folder As String
    folder = ThisWorkbook.Path & "\vba\"
    If Len(Dir(folder, vbDirectory)) = 0 Then
        MsgBox "vba folder not found: " & folder, vbExclamation
        Exit Sub
    End If

    ' Collect every standard module (1) and class module (2). Document modules
    ' (100, e.g. ThisWorkbook / sheets) are left alone.
    Dim comp As Object, toRemove As Collection
    Set toRemove = New Collection
    For Each comp In proj.VBComponents
        If comp.Type = 1 Or comp.Type = 2 Then toRemove.Add comp
    Next comp
    Dim c As Object
    For Each c In toRemove
        proj.VBComponents.Remove c
    Next c

    ' Import the canonical set, in dependency-friendly order.
    Dim files As Variant, i As Long, n As Long
    files = Array("OrderedDict.cls", "JsonParser.bas", "JsonWriter.bas", _
                  "CobolEncoding.bas", "CobolParser.bas", "CobolLogicViewer.bas", _
                  "Main.bas", "TestRunner.bas", _
                  "Test_Phase1.bas", "Test_Phase2.bas", "Test_Phase3.bas", _
                  "Test_Phase4.bas", "Test_Phase5.bas")
    n = 0
    For i = LBound(files) To UBound(files)
        If Len(Dir(folder & files(i))) > 0 Then
            proj.VBComponents.Import folder & files(i)
            n = n + 1
        Else
            Debug.Print "MISSING: " & folder & files(i)
        End If
    Next i

    MsgBox "Rebuilt " & n & " modules from " & folder, vbInformation
    Exit Sub

NoTrust:
    MsgBox "Cannot access the VBA project." & vbCrLf & _
           "Enable: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
           "Macro Settings > 'Trust access to the VBA project object model'.", vbExclamation
End Sub
