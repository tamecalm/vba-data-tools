Sub FixCompoundDirections()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim regex1 As Object, regex2 As Object

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' Fix 1: "NORTH-E EAST" → "NORTH-EAST"
    Set regex1 = CreateObject("VBScript.RegExp")
    With regex1
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b(NORTH|SOUTH)-(E|W)\s+EAST\b"
    End With

    ' Fix 2: "NORTH-EEAST" → "NORTH-EAST"
    Set regex2 = CreateObject("VBScript.RegExp")
    With regex2
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b(NORTH|SOUTH)-([EW])EAST\b"
    End With

    For i = 1 To lastRow
        If Not IsError(ws.Cells(i, "F").Value) Then
            cellValue = CStr(ws.Cells(i, "F").Value)

            ' Apply both fixes
            If regex1.test(cellValue) Then
                cellValue = regex1.Replace(cellValue, "$1-$2AST")
            End If
            If regex2.test(cellValue) Then
                cellValue = regex2.Replace(cellValue, "$1-$2AST")
            End If

            ws.Cells(i, "F").Value = cellValue
        End If
    Next i

    MsgBox "Compound direction corrections completed in column F.", vbInformation
End Sub
