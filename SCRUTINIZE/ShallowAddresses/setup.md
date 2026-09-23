# 📘 VBA Macro Setup Guide: ShallowAddressesReport

## ✨ Overview

In Nigerian address verification, an address like `"18 SARI STREET"` or `"AIRPORT ROAD"` cannot be verified by a field officer or rider because it lacks **location depth** (no town, area/district, LGA, landmark, or bus stop).

This macro identifies addresses that are **at most 3 words long** (`Word Count <= 3`), preventing compound addresses (such as `"NO 2 AKINBOREWA CLOSE ADEEKO OLUSANYA STREET"`) from being falsely flagged.

Furthermore, it generates two coordinated sheets:
1. **`ShallowAddresses`**: Contains the full audit list of all flagged addresses with a Pivot Table summary (`ShallowPivot`).
2. **`FinalRecord`**: A dedicated sheet created specifically for logging into your **Master Record Table**, containing:
   - **Summary Table (Columns A:D)**: Lists each officer, `Total Found`, `50% Deduction (Round Up)`, and `25% Deduction from 50% (Round Up)` using dynamic Excel formulas.
   - **50% Raw Address Records (Columns F:L)**: Automatically extracts the 50% raw address rows for each officer, indicating which ones fall within the net 25% deduction (`Yes` / `No (25% Relieved)`), linked via Excel formulas and clickable jump hyperlinks back to `ShallowAddresses`.
   - **Final Record Pivot Table (Columns N:R)**: A dedicated Pivot Table (`FinalRecordPivot`) summarizing the 50% selected records by **Verification Officer** and **Incomplete Category**.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- In Excel, press `ALT + F11` to launch the **Visual Basic for Applications (VBA)** editor.

### Step 2: Insert a Module

- In the left panel (**Project Explorer**), right-click your workbook name.
- Choose: **Insert > Module**

### Step 3: Paste the Macro Code

Paste the code below:

```vba
' === Macro: ShallowAddressesReport ===
' Description: Detects shallow/incomplete addresses strictly up to 3 words,
'              builds a full audit sheet ("ShallowAddresses"), and generates
'              a dedicated "FinalRecord" sheet with 50% and 25%-from-50% rounded-up formulas,
'              linked raw address rows, and a dedicated Pivot Table.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 2025

Sub ShallowAddressesReport()
    Dim ws As Worksheet, wsNew As Worksheet, wsRecord As Worksheet
    Dim lastRow As Long, newRow As Long, rowIdx As Long, recRow As Long, detailRow As Long
    Dim ptStart As Range, dataRange As Range
    Dim ptRecordStart As Range, recordDataRange As Range
    Dim officer As Variant
    Dim rawAddr As String, cleanAddr As String, category As String
    Dim addressCol As String, officerCol As String
    Dim wc As Long, totalCount As Long, takeCount50 As Long, takeCount25 As Long, i As Long, srcRow As Long
    Dim regexSuffix As Object
    Dim ptCache As PivotCache, pt As PivotTable
    Dim ptRecordCache As PivotCache, ptRecord As PivotTable
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
        
        ' Create FinalRecord sheet
        Set wsRecord = ThisWorkbook.Sheets.Add(After:=wsNew)
        wsRecord.Name = "FinalRecord"
        
        With wsRecord
            .Range("A1").Value = "Verification Officer"
            .Range("B1").Value = "Total Found"
            .Range("C1").Value = "50% Deduction (Round Up)"
            .Range("D1").Value = "25% Deduction from 50% (Round Up)"
            .Range("A1:D1").Font.Bold = True
            
            .Range("F1").Value = "Address (50% Selection)"
            .Range("G1").Value = "Verification Officer"
            .Range("H1").Value = "Incomplete Category"
            .Range("I1").Value = "Word Count"
            .Range("J1").Value = "Source Sheet Row"
            .Range("K1").Value = "Within 25% Net?"
            .Range("L1").Value = "Link to Shallow Sheet"
            .Range("F1:L1").Font.Bold = True
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
            wsRecord.Cells(recRow, 3).Formula = "=ROUNDUP(B" & recRow & "*0.5, 0)"
            wsRecord.Cells(recRow, 4).Formula = "=ROUNDUP(C" & recRow & "*0.75, 0)"
            
            rowsList = Split(dictOfficers(offKey), ",")
            totalCount = UBound(rowsList) - LBound(rowsList) + 1
            takeCount50 = Application.WorksheetFunction.RoundUp(totalCount * 0.5, 0)
            takeCount25 = Application.WorksheetFunction.RoundUp(takeCount50 * 0.75, 0)
            
            For i = 0 To takeCount50 - 1
                srcRow = CLng(rowsList(i))
                wsRecord.Cells(detailRow, 6).Formula = "='ShallowAddresses'!A" & srcRow
                wsRecord.Cells(detailRow, 7).Formula = "='ShallowAddresses'!B" & srcRow
                wsRecord.Cells(detailRow, 8).Formula = "='ShallowAddresses'!C" & srcRow
                wsRecord.Cells(detailRow, 9).Formula = "='ShallowAddresses'!D" & srcRow
                wsRecord.Cells(detailRow, 10).Formula = "='ShallowAddresses'!E" & srcRow
                If i < takeCount25 Then
                    wsRecord.Cells(detailRow, 11).Value = "Yes"
                Else
                    wsRecord.Cells(detailRow, 11).Value = "No (25% Relieved)"
                End If
                wsRecord.Cells(detailRow, 12).Formula = "=HYPERLINK(""#'ShallowAddresses'!A" & srcRow & """, ""Go to Row " & srcRow & """)"
                detailRow = detailRow + 1
            Next i
            
            recRow = recRow + 1
        Next offKey
        
        wsRecord.Cells(recRow, 1).Value = "Grand Total"
        wsRecord.Cells(recRow, 2).Formula = "=SUM(B2:B" & (recRow - 1) & ")"
        wsRecord.Cells(recRow, 3).Formula = "=SUM(C2:C" & (recRow - 1) & ")"
        wsRecord.Cells(recRow, 4).Formula = "=SUM(D2:D" & (recRow - 1) & ")"
        wsRecord.Range("A" & recRow & ":D" & recRow).Font.Bold = True
        
        ' Create Pivot Table for FinalRecord
        If detailRow > 2 Then
            Set recordDataRange = wsRecord.Range("F1:J" & (detailRow - 1))
            Set ptRecordStart = wsRecord.Range("N3")
            
            Set ptRecordCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=recordDataRange)
            
            On Error Resume Next
            Set ptRecord = wsRecord.PivotTables("FinalRecordPivot")
            On Error GoTo 0
            
            If ptRecord Is Nothing Then
                Set ptRecord = ptRecordCache.CreatePivotTable(TableDestination:=ptRecordStart, TableName:="FinalRecordPivot")
            End If
            
            With ptRecord
                .ClearAllFilters
                .PivotFields("Verification Officer").Orientation = xlRowField
                .PivotFields("Incomplete Category").Orientation = xlColumnField
                .AddDataField .PivotFields("Address (50% Selection)"), "Count of 50% Recorded", xlCount
            End With
        End If
        
        wsRecord.Columns("A:R").AutoFit
        
        wsRecord.Activate
        MsgBox (newRow - 2) & " shallow addresses flagged." & vbCrLf & _
               "Pivot Table generated in 'ShallowAddresses'." & vbCrLf & _
               "Final Record Table (50% & 25% from 50%), raw rows & Pivot Table generated in 'FinalRecord'.", vbInformation
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
```

---

## ✅ How to Run the Macro

1. Save your workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
2. Press `ALT + F8`, select `ShallowAddressesReport`, and click **Run**.
3. When prompted:
   - Enter the column letter containing addresses (default: `C`).
   - Enter the column letter containing officer names (default: `L`).
4. Two sheets will be generated:
   - **`ShallowAddresses`**: Complete list of flagged addresses and Pivot Table summary (`ShallowPivot`).
   - **`FinalRecord`**: The summary table (with both 50% and 25%-from-50% deductions), 50% raw address rows, and the Pivot Table (`FinalRecordPivot`).

---

## 🧪 Structure of `FinalRecord` Sheet

### Summary Record Table (Columns A:D)
| Verification Officer | Total Found | 50% Deduction (Round Up) | 25% Deduction from 50% (Round Up) |
| :--- | :---: | :---: | :---: |
| Officer A | 10 | **5** | **4** *(5 − 25% of 5 = 3.75 $\rightarrow$ 4)* |
| Officer B | 7 | **4** | **3** *(4 − 25% of 4 = 3)* |
| Officer C | 2 | **1** | **1** |
| **Grand Total** | **19** | **10** | **8** |

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)
- **Date Created:** July 2025
- **Workbook Requirement:** Excel Macro-Enabled Workbook (`.xlsm`)
