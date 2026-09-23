# 📘 VBA Macro Setup Guide: IncompleteAddressesReport

## ✨ Overview

This macro automates the detection of **incomplete addresses** in an active worksheet. An address is flagged as incomplete if it either:
- Contains **no numbers** (e.g., text-only entries such as `"Broad Street"`)
- Contains **no letters** (e.g., numbers-only entries such as `"12345"`)

The macro extracts each incomplete address, maps it to its assigned **verification officer**, classifies the address type, and dynamically builds a **Pivot Table summary**.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- In Excel, press `ALT + F11` to launch the **Visual Basic for Applications (VBA)** editor.

### Step 2: Insert a Module

- In the left panel (**Project Explorer**), right-click on your workbook name.
- Choose: **Insert > Module**

### Step 3: Paste the Macro Code

Paste the code below:

```vba
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
```

---

## ✅ How to Run the Macro

1. Save your workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.
2. Press `ALT + F8`, select `IncompleteAddressesReport`, and click **Run**.
3. When prompted:
   - Enter the column letter for addresses (default: `C`).
   - Enter the column letter for verification officer names (default: `L`).
4. The macro will create a new sheet named `IncompleteAddresses` containing the detailed breakdown and a pivot table.

---

## 🧪 Example

### Sample Input Data:

| C (Address)       | L (Officer)  |
|-------------------|--------------|
| "Broad Street"    | "Officer A"  |
| "12 Palm Avenue"  | "Officer B"  |
| "98765"           | "Officer C"  |
| "Victoria Island" |              |

### Output in `IncompleteAddresses` Sheet:

| Incomplete Address | Verification Officer | Address Type  |
|--------------------|----------------------|---------------|
| Broad Street       | Officer A            | Text-only     |
| 98765              | Officer C            | Numbers-only  |
| Victoria Island    | Unassigned           | Text-only     |

*(Note: `"12 Palm Avenue"` is excluded as it contains both letters and numbers).*

A pivot table is placed starting at cell `E3`, summarizing the count of incomplete addresses by verification officer and address type.

---

## 📝 Notes

- Uses `VBScript.RegExp` for pattern matching without requiring manual character loops.
- Automatically handles unassigned officers by tagging them as `"Unassigned"`.
- Existing `IncompleteAddresses` sheets are replaced automatically on each run.
- Canceling either column input prompt safely cancels execution.

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)
- **Date Created:** July 2025
- **Workbook Requirement:** Macro-Enabled Workbook (`.xlsm`)
