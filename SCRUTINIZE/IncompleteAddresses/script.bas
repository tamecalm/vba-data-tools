' === Macro: IncompleteAddressesReport ===
' Description: Scans an address column for incomplete addresses (text-only or numbers-only),
'              extracts them alongside assigned verification officers, and creates a pivot summary.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 2025

Sub IncompleteAddressesReport()
    Dim ws As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, newRow As Long
    Dim cell As Range
    Dim officer As Variant
    Dim address As String, addressType As String
    Dim addressCol As String, officerCol As String
    Dim regexDigit As Object, regexLetter As Object
    Dim ptCache As PivotCache, pt As PivotTable
    Dim ptStart As Range, dataRange As Range
    
    addressCol = InputBox("Enter the column letter that contains addresses:", "Address Column", "C")
    If Trim(addressCol) = "" Then Exit Sub
    
    officerCol = InputBox("Enter the column letter that contains officer names:", "Officer Column", "L")
    If Trim(officerCol) = "" Then Exit Sub
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, addressCol).End(xlUp).Row
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsNew = ThisWorkbook.Sheets("IncompleteAddresses")
    If Not wsNew Is Nothing Then wsNew.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsNew = ThisWorkbook.Sheets.Add
    wsNew.Name = "IncompleteAddresses"
    
    With wsNew
        .Range("A1").Value = "Incomplete Address"
        .Range("B1").Value = "Verification Officer"
        .Range("C1").Value = "Address Type"
        .Range("A1:C1").Font.Bold = True
    End With
    
    newRow = 2
    
    Set regexDigit = CreateObject("VBScript.RegExp")
    regexDigit.Pattern = "\d"
    regexDigit.IgnoreCase = True
    regexDigit.Global = True
    
    Set regexLetter = CreateObject("VBScript.RegExp")
    regexLetter.Pattern = "[A-Za-z]"
    regexLetter.IgnoreCase = True
    regexLetter.Global = True
    
    For Each cell In ws.Range(addressCol & "2:" & addressCol & lastRow)
        If Not IsEmpty(cell.Value) Then
            address = Trim(cell.Value)
            
            If (Not regexDigit.Test(address)) Or (Not regexLetter.Test(address)) Then
                If Not regexDigit.Test(address) Then
                    addressType = "Text-only"
                ElseIf Not regexLetter.Test(address) Then
                    addressType = "Numbers-only"
                Else
                    addressType = "Other"
                End If
                
                officer = Trim(ws.Cells(cell.Row, officerCol).Value)
                If officer = "" Then officer = "Unassigned"
                
                wsNew.Cells(newRow, 1).Value = address
                wsNew.Cells(newRow, 2).Value = officer
                wsNew.Cells(newRow, 3).Value = addressType
                newRow = newRow + 1
            End If
        End If
    Next cell
    
    wsNew.Columns("A:C").AutoFit
    
    If newRow > 2 Then
        Set dataRange = wsNew.Range("A1").CurrentRegion
        Set ptStart = wsNew.Range("E3")
        
        Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
        
        On Error Resume Next
        Set pt = wsNew.PivotTables("IncompletePivot")
        On Error GoTo 0
        
        If pt Is Nothing Then
            Set pt = ptCache.CreatePivotTable(TableDestination:=ptStart, TableName:="IncompletePivot")
        End If
        
        With pt
            .ClearAllFilters
            .PivotFields("Verification Officer").Orientation = xlRowField
            .PivotFields("Address Type").Orientation = xlColumnField
            .AddDataField .PivotFields("Incomplete Address"), "Count of Incomplete Address", xlCount
        End With
    End If
    
    MsgBox "Incomplete addresses report generated with Pivot Table in 'IncompleteAddresses'.", vbInformation
End Sub
