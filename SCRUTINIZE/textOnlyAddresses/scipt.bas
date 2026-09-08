Sub CountTextOnlyAddressesByOfficer()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim officer As Variant
    Dim i As Long
    Dim isTextOnly As Boolean
    Dim newRow As Long
   
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
   
    ' Find the last row in column C
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
   
    ' Create a new worksheet to store the results
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsNew = ThisWorkbook.Sheets("TextOnlyAddresses")
    If Not wsNew Is Nothing Then wsNew.Delete ' Delete if it already exists
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsNew = ThisWorkbook.Sheets.Add
    wsNew.Name = "TextOnlyAddresses"
   
    ' Add headers to the new sheet
    wsNew.Range("A1").Value = "Address"
    wsNew.Range("B1").Value = "Verification Officer"
   
    ' Start populating from row 2 in the new sheet
    newRow = 2
   
    ' Loop through each cell in column C
    For Each cell In ws.Range("C1:C" & lastRow)
        If Not IsEmpty(cell.Value) Then
            ' Get the officer assigned to the address
            officer = ws.Cells(cell.Row, "L").Value
           
            ' Check if the address contains any numbers
            isTextOnly = True
            For i = 0 To 9
                If InStr(1, cell.Value, CStr(i)) > 0 Then
                    isTextOnly = False
                    Exit For
                End If
            Next i
           
            ' If the address is text-only, add it to the new sheet
            If isTextOnly Then
                wsNew.Cells(newRow, 1).Value = cell.Value ' Address in column A
                wsNew.Cells(newRow, 2).Value = officer    ' Verification Officer in column B
                newRow = newRow + 1
            End If
        End If
    Next cell
   
    MsgBox "Text-only addresses have been listed in the new sheet: 'TextOnlyAddresses'.", vbInformation
End Sub