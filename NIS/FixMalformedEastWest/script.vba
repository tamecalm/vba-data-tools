'---------------------------------------------------------------------------------------
' Module      : FixMalformedEastWest
' Author      : Ilesanmi Kehinde John (aka Calm) 🧘‍♂️
' Created     : July 29, 2025
' Description : Fixes malformed compound directions in Excel column F such as:
'               - "E EAST" → "EAST"
'               - "W EAST" → "WEST"
'               It scans each row of column F in the active worksheet and replaces
'               incorrect direction notations with valid ones.
'---------------------------------------------------------------------------------------

Sub FixMalformedEastWest()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim updatedValue As String
    Dim pos As Long
    Dim countFixes As Long

    ' Reference to the active sheet
    Set ws = ActiveSheet

    ' Find the last used row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    countFixes = 0

    ' Loop through each cell in column F
    For i = 1 To lastRow
        If Not IsError(ws.Cells(i, "F").Value) Then
            cellValue = Trim(CStr(ws.Cells(i, "F").Value))
            updatedValue = cellValue

            ' Fix "E EAST" → "EAST"
            pos = InStr(1, UCase(updatedValue), "E EAST")
            If pos > 0 Then
                updatedValue = Left(updatedValue, pos - 1) & "EAST"
                countFixes = countFixes + 1
            End If

            ' Fix "W EAST" → "WEST"
            pos = InStr(1, UCase(updatedValue), "W EAST")
            If pos > 0 Then
                updatedValue = Left(updatedValue, pos - 1) & "WEST"
                countFixes = countFixes + 1
            End If

            ' Update cell if changes were made
            If updatedValue <> cellValue Then
                ws.Cells(i, "F").Value = updatedValue
            End If
        End If
    Next i

    ' Notify user of completion and how many rows were affected
    MsgBox "Malformed 'E EAST' and 'W EAST' corrections completed in column F." & vbCrLf & _
           "Total fixed entries: " & countFixes, vbInformation
End Sub
