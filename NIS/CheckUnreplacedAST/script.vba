Sub CheckRemainingASTOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim foundCount As Long
    Dim regex As Object

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    foundCount = 0

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "AST\b" ' Match anything ending with AST
    End With

    ' Clear previous highlights
    ws.Range("F1:F" & lastRow).Interior.ColorIndex = xlNone

    For i = 1 To lastRow
        cellValue = ws.Cells(i, "F").Value

        If Not IsError(cellValue) Then
            cellText = Trim(CStr(cellValue))

            ' Highlight only if it ends in AST but not EAST
            If regex.test(cellText) And LCase(Right(cellText, 4)) <> "east" Then
                ws.Cells(i, "F").Interior.Color = vbYellow
                foundCount = foundCount + 1
            End If
        End If
    Next i

    MsgBox foundCount & " cell(s) still end in 'AST' (not 'EAST').", vbInformation
End Sub
