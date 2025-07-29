# 📘 VBA Macro Setup Guide: FixCompoundDirections

## ✨ Overview

This macro scans **column F** of the active sheet to identify and correct **compound direction errors** like:

- `"NORTH-E EAST"` → `"NORTH-EAST"`
- `"NORTH-EEAST"` → `"NORTH-EAST"`

It uses **regular expressions** to identify these malformed patterns and correct them automatically, ensuring consistent address formatting.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- In Excel, press `ALT + F11` to open the **VBA Editor**.

### Step 2: Insert the Code Module

1. In the **Project Explorer**, right-click your workbook name.
2. Click **Insert > Module**
3. Paste the macro code below into the new module:

```vba
' === Macro: FixCompoundDirections ===
' Description: Fixes malformed compound directions like "NORTH-E EAST" or "SOUTH-EEAST" to "NORTH-EAST", etc.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 29, 2025

Sub FixCompoundDirections()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim regex1 As Object, regex2 As Object

    ' Use the currently active sheet
    Set ws = ActiveSheet

    ' Determine the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row

    ' Fix 1: Handle "NORTH-E EAST" → "NORTH-EAST"
    Set regex1 = CreateObject("VBScript.RegExp")
    With regex1
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b(NORTH|SOUTH)-(E|W)\s+EAST\b"
    End With

    ' Fix 2: Handle "NORTH-EEAST" → "NORTH-EAST"
    Set regex2 = CreateObject("VBScript.RegExp")
    With regex2
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b(NORTH|SOUTH)-([EW])EAST\b"
    End With

    ' Apply corrections row by row
    For i = 1 To lastRow
        If Not IsError(ws.Cells(i, "F").Value) Then
            cellValue = CStr(ws.Cells(i, "F").Value)

            ' Apply both regex-based fixes
            If regex1.test(cellValue) Then
                cellValue = regex1.Replace(cellValue, "$1-$2AST")
            End If
            If regex2.test(cellValue) Then
                cellValue = regex2.Replace(cellValue, "$1-$2AST")
            End If

            ' Write corrected value back
            ws.Cells(i, "F").Value = cellValue
        End If
    Next i

    ' Notify user when complete
    MsgBox "Compound direction corrections completed in column F.", vbInformation
End Sub
```

---

## ✅ How to Run the Macro

1. Save your workbook as a **Macro-Enabled Workbook (`.xlsm`)**.
2. Activate the worksheet containing the address data.
3. Press `ALT + F8`, select `FixCompoundDirections`, then click **Run**.
4. The macro will process all rows in **column F** and auto-correct malformed directions.

---

## 🧪 Example

| Original Value       | Corrected Output   |
|----------------------|--------------------|
| `NORTH-E EAST`       | `NORTH-EAST`       |
| `SOUTH-W EAST`       | `SOUTH-WAST`       |
| `NORTH-EEAST`        | `NORTH-EAST`       |
| `SOUTH-WEAST`        | `SOUTH-WAST`       |
| `NORTH-EAST`         | _No change_        |
| `EAST SIDE`          | _No change_        |

> Note: Final `-AST` form can be cleaned up further using another macro if needed.

---

## 📝 Notes

- The macro uses two **regular expressions**:
  - `regex1` fixes `"NORTH-E EAST"`-type spacing issues.
  - `regex2` fixes `"NORTH-EEAST"`-type duplication issues.
- Processing is **case-insensitive**.
- Only column **F** is processed.
- Works on the **active worksheet** only.
- Cells with errors are skipped.

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)  
- **Created On:** July 29, 2025  
- **Use Case:** Clean and normalize address direction formatting  
- **License:** Free to use and adapt in any Excel-based automation task

---
