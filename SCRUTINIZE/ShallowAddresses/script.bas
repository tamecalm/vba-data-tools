' === Macro: ShallowAddressesReport ===
' Description: Detects shallow/incomplete addresses strictly up to 3 words,
'              builds a full audit sheet ("ShallowAddresses"), and generates
'              a dedicated "FinalRecord" sheet with 50% rounded-up formulas
'              and the corresponding 50% raw address rows linked back to the source.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 2025

Sub ShallowAddressesReport()
    Dim ws As Worksheet, wsNew As Worksheet, wsRecord As Worksheet
    Dim lastRow As Long, newRow As Long, rowIdx As Long, recRow As Long, detailRow As Long
    Dim ptStart As Range, dataRange As Range
    Dim officer As Variant
    Dim rawAddr As String, cleanAddr As String, category As String
    Dim addressCol As String, officerCol As String
    Dim wc As Long, totalCount As Long, takeCount As Long, i As Long, srcRow As Long
    Dim regexSuffix As Object
    Dim ptCache As PivotCache, pt As PivotTable
    Dim dictOfficers As Object, offKey As Variant, rowsList() As String
    
    addressCol = InputBox("Enter the column letter that contains addresses:", "Address Column", "C")
    If Trim(addressCol) = "" Then Exit Sub
    
    officerCol = InputBox("Enter the column letter that contains officer names:", "Officer Column", "L")
    If Trim(officerCol) = "" Then Exit Sub
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, addressCol).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data found in column " & addressCol, vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsRecord = ThisWorkbook.Sheets("FinalRecord")
    If Not wsRecord Is Nothing Then wsRecord.Delete
    Set wsNew = ThisWorkbook.Sheets("ShallowAddresses")
    If Not wsNew Is Nothing Then wsNew.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ws)
    wsNew.Name = "ShallowAddresses"
    
    With wsNew
        .Range("A1").Value = "Address"
        .Range("B1").Value = "Verification Officer"
        .Range("C1").Value = "Incomplete Category"
        .Range("D1").Value = "Word Count"
        .Range("E1").Value = "Source Row"
        .Range("A1:E1").Font.Bold = True
    End With
    
    newRow = 2
    
    Set regexSuffix = CreateObject("VBScript.RegExp")
    regexSuffix.Pattern = "(STREET|ST|ROAD|RD|AVENUE|AVE|CLOSE|CLS|CRESCENT|CRES|WAY|LANE|LN|DRIVE|DR|BOULEVARD|BLVD|HIGHWAY|EXPRESSWAY|EXP|LINE)\.?$"
    regexSuffix.IgnoreCase = True
    regexSuffix.Global = False
    
    For rowIdx = 2 To lastRow
        rawAddr = Trim(CStr(ws.Cells(rowIdx, addressCol).Value))
        
        If rawAddr <> "" Then
            cleanAddr = CleanAddressText(rawAddr)
            wc = CountWords(cleanAddr)
            category = ""
            
            If wc = 1 Then
                category = "Single-word"
            ElseIf wc = 2 Then
                category = "Two-word"
            ElseIf wc = 3 Then
                If regexSuffix.Test(cleanAddr) Then
                    category = "Bare Street (3 words)"
                Else
                    category = "Three-word Address"
                End If
            End If
            
            If category <> "" Then
                officer = Trim(CStr(ws.Cells(rowIdx, officerCol).Value))
                If officer = "" Then officer = "Unassigned"
                
                wsNew.Cells(newRow, 1).Value = rawAddr
                wsNew.Cells(newRow, 2).Value = officer
                wsNew.Cells(newRow, 3).Value = category
                wsNew.Cells(newRow, 4).Value = wc
                wsNew.Cells(newRow, 5).Value = rowIdx
                newRow = newRow + 1
            End If
        End If
    Next rowIdx
    
    wsNew.Columns("A:E").AutoFit
    
    If newRow > 2 Then
        Set dataRange = wsNew.Range("A1:E" & (newRow - 1))
        Set ptStart = wsNew.Range("G3")
        
        Set ptCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)
        
        On Error Resume Next
        Set pt = wsNew.PivotTables("ShallowPivot")
        On Error GoTo 0
        
        If pt Is Nothing Then
            Set pt = ptCache.CreatePivotTable(TableDestination:=ptStart, TableName:="ShallowPivot")
        End If
        
        With pt
            .ClearAllFilters
            .PivotFields("Verification Officer").Orientation = xlRowField
            .PivotFields("Incomplete Category").Orientation = xlColumnField
            .AddDataField .PivotFields("Address"), "Count of Incomplete", xlCount
        End With
        
        ' Create FinalRecord sheet linking to ShallowAddresses
        Set wsRecord = ThisWorkbook.Sheets.Add(After:=wsNew)
        wsRecord.Name = "FinalRecord"
        
        With wsRecord
            .Range("A1").Value = "Verification Officer"
            .Range("B1").Value = "Total Found"
            .Range("C1").Value = "50% Value"
            .Range("D1").Value = "Final to Record (50% Round Up)"
            .Range("A1:D1").Font.Bold = True
            
            .Range("F1").Value = "Address (50% Selection)"
            .Range("G1").Value = "Verification Officer"
            .Range("H1").Value = "Incomplete Category"
            .Range("I1").Value = "Word Count"
            .Range("J1").Value = "Source Sheet Row"
            .Range("K1").Value = "Link to Shallow Sheet"
            .Range("F1:K1").Font.Bold = True
        End With
        
        Set dictOfficers = CreateObject("Scripting.Dictionary")
        For rowIdx = 2 To newRow - 1
            officer = Trim(CStr(wsNew.Cells(rowIdx, 2).Value))
            If Not dictOfficers.Exists(officer) Then
                dictOfficers.Add officer, ""
            End If
            If dictOfficers(officer) = "" Then
                dictOfficers(officer) = CStr(rowIdx)
            Else
                dictOfficers(officer) = dictOfficers(officer) & "," & CStr(rowIdx)
            End If
        Next rowIdx
        
        recRow = 2
        detailRow = 2
        
        For Each offKey In dictOfficers.Keys
            wsRecord.Cells(recRow, 1).Value = offKey
            wsRecord.Cells(recRow, 2).Formula = "=COUNTIF('ShallowAddresses'!B:B, A" & recRow & ")"
            wsRecord.Cells(recRow, 3).Formula = "=B" & recRow & "*0.5"
            wsRecord.Cells(recRow, 4).Formula = "=ROUNDUP(B" & recRow & "*0.5, 0)"
            
            rowsList = Split(dictOfficers(offKey), ",")
            totalCount = UBound(rowsList) - LBound(rowsList) + 1
            takeCount = Application.WorksheetFunction.RoundUp(totalCount * 0.5, 0)
            
            For i = 0 To takeCount - 1
                srcRow = CLng(rowsList(i))
                wsRecord.Cells(detailRow, 6).Formula = "='ShallowAddresses'!A" & srcRow
                wsRecord.Cells(detailRow, 7).Formula = "='ShallowAddresses'!B" & srcRow
                wsRecord.Cells(detailRow, 8).Formula = "='ShallowAddresses'!C" & srcRow
                wsRecord.Cells(detailRow, 9).Formula = "='ShallowAddresses'!D" & srcRow
                wsRecord.Cells(detailRow, 10).Formula = "='ShallowAddresses'!E" & srcRow
                wsRecord.Cells(detailRow, 11).Formula = "=HYPERLINK(""#'ShallowAddresses'!A" & srcRow & """, ""Go to Row " & srcRow & """)"
                detailRow = detailRow + 1
            Next i
            
            recRow = recRow + 1
        Next offKey
        
        wsRecord.Cells(recRow, 1).Value = "Grand Total"
        wsRecord.Cells(recRow, 2).Formula = "=SUM(B2:B" & (recRow - 1) & ")"
        wsRecord.Cells(recRow, 3).Formula = "=SUM(C2:C" & (recRow - 1) & ")"
        wsRecord.Cells(recRow, 4).Formula = "=SUM(D2:D" & (recRow - 1) & ")"
        wsRecord.Range("A" & recRow & ":D" & recRow).Font.Bold = True
        
        wsRecord.Columns("A:K").AutoFit
        
        wsRecord.Activate
        MsgBox (newRow - 2) & " shallow addresses flagged." & vbCrLf & _
               "Pivot Table generated in 'ShallowAddresses'." & vbCrLf & _
               "50% Final Record Table & raw rows generated in 'FinalRecord'.", vbInformation
    Else
        MsgBox "No shallow or incomplete addresses (3 words or fewer) found in column " & addressCol & ".", vbInformation
    End If
End Sub

Private Function CountWords(ByVal txt As String) As Long
    txt = Trim(txt)
    If Len(txt) = 0 Then
        CountWords = 0
        Exit Function
    End If
    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    Dim parts() As String
    parts = Split(txt, " ")
    CountWords = UBound(parts) - LBound(parts) + 1
End Function

Private Function CleanAddressText(ByVal txt As String) As String
    txt = Trim(txt)
    Do While Len(txt) > 0 And (Right(txt, 1) = "," Or Right(txt, 1) = "." Or Right(txt, 1) = "-" Or Right(txt, 1) = ";")
        txt = Trim(Left(txt, Len(txt) - 1))
    Loop
    CleanAddressText = txt
End Function
