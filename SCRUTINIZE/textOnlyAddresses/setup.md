# 📘 VBA Macro Setup Guide: CountTextOnlyAddressesByOfficer

## ✨ Overview

This macro scans column **C** of the active sheet for addresses that **do not contain any numbers** (i.e., are strictly text-only), then creates a new worksheet listing:
- Each valid **text-only address**
- The corresponding **verification officer** from column **L**

This is helpful for separating non-numeric address entries by assigned officers for verification or reporting purposes.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- In Excel, press `ALT + F11` to launch the **Visual Basic for Applications (VBA)** editor.

### Step 2: Insert a Module

- In the left panel (**Project Explorer**), right-click on your workbook name.
- Choose: **Insert > Module**

### Step 3: Paste the Macro Code

Paste the code below exactly as it is:

```vba
' === Macro: CountTextOnlyAddressesByOfficer ===
' Description: Extracts addresses with no digits from column C, along with their verification officers from column L, into a new sheet.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 29, 2025

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
```

---

## ✅ How to Run the Macro

1. Save your workbook as a **Macro-Enabled Workbook (`.xlsm`)**.
2. Press `ALT + F8`, select `CountTextOnlyAddressesByOfficer`, then click **Run**.
3. Ensure your worksheet:
   - Has addresses in **Column C**
   - Has corresponding officers in **Column L**

---

## 🧪 Example

Given this sample data:

| C (Address)    | L (Officer)      |
|----------------|------------------|
| "Broad Street" | "Officer A"      |
| "45 George St" | "Officer B"      |
| "Kingsway"     | "Officer C"      |

After running the macro, a new sheet `TextOnlyAddresses` will contain:

| Address       | Verification Officer |
|---------------|----------------------|
| Broad Street  | Officer A            |
| Kingsway      | Officer C            |

> `"45 George St"` is excluded because it contains digits.

---

## 📝 Notes

- The macro ignores **empty rows** in column C.
- If a sheet named **TextOnlyAddresses** exists, it will be deleted and recreated.
- The officer’s name is pulled from the same row in column **L**.

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)  
- **Created On:** November 25, 2024   
- **Workbook Requirement:** Enable macros and allow programmatic access  
- **License:** Free to use and adapt

---
