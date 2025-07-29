Sub ReplaceASTWithEast()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim regex As Object

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .Pattern = "(\S*)AST\b" ' Match non-space characters ending in AST
    End With

    For i = 1 To lastRow
        If Not IsError(ws.Cells(i, "F").Value) Then
            cellValue = CStr(ws.Cells(i, "F").Value)
            If regex.test(cellValue) Then
                ws.Cells(i, "F").Value = regex.Replace(cellValue, "$1 EAST")
            End If
        End If
    Next i

    MsgBox "AST replacements completed in column F.", vbInformation
End Sub
