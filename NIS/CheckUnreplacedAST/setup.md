# 📘 VBA Macro Setup Guide: CheckRemainingASTOnly

## ✨ Overview

This macro scans **column F** of the active sheet to detect and highlight any cells that:

- **End in `AST`**, **excluding** those that end in `"EAST"`
- It **highlights them in yellow**
- Finally, it reports how many such cells were found

This is useful for validating data cleanup where `"AST"` is being corrected to `"EAST"` and ensures no leftover mismatches remain.

---

## 🔧 Setup Instructions

### Step 1: Open the VBA Editor

- Open Excel and press `ALT + F11` to launch the **VBA editor**.

### Step 2: Insert the Code

1. In the **Project Explorer**, right-click your workbook.
2. Choose: **Insert > Module**
3. Paste the macro code below into the module:

```vba
' === Macro: CheckRemainingASTOnly ===
' Description: Highlights cells in column F that end with 'AST' but not 'EAST', using regular expression pattern matching.
' Author: Ilesanmi Kehinde John (Calm)
' Date Created: July 29, 2025

Sub CheckRemainingASTOnly()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As Variant
    Dim cellText As String
    Dim foundCount As Long
    Dim regex As Object

    ' Target the active worksheet
    Set ws = ActiveSheet

    ' Get the last row in column F
    lastRow = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    foundCount = 0

    ' Set up RegExp to detect words ending in 'AST'
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Pattern = "AST\b"
    End With

    ' Clear previous highlights
    ws.Range("F1:F" & lastRow).Interior.ColorIndex = xlNone

    ' Loop through each cell in column F
    For i = 1 To lastRow
        cellValue = ws.Cells(i, "F").Value

        If Not IsError(cellValue) Then
            cellText = Trim(CStr(cellValue))

            ' Highlight if ends in 'AST' but not 'EAST'
            If regex.test(cellText) And LCase(Right(cellText, 4)) <> "east" Then
                ws.Cells(i, "F").Interior.Color = vbYellow
                foundCount = foundCount + 1
            End If
        End If
    Next i

    ' Show count of remaining ASTs
    MsgBox foundCount & " cell(s) still end in 'AST' (not 'EAST').", vbInformation
End Sub
```

---

## ✅ How to Run the Macro

1. Save your workbook as a **Macro-Enabled Workbook (`.xlsm`)**.
2. Press `ALT + F8`, select `CheckRemainingASTOnly`, then click **Run**.
3. Make sure you’re on the worksheet with the target data in **column F**.

---

## 🧪 Example

| Column F Text      | Highlighted? |
|--------------------|--------------|
| "SOUTHWEST AST"    | ✅ Yes       |
| "SOUTHEAST EAST"   | ❌ No        |
| "AST"              | ✅ Yes       |
| "SPEEDAST"         | ✅ Yes       |
| "WEST COAST"       | ❌ No        |

- Any text **ending in "AST"** but **not "EAST"** is highlighted in **yellow**.
- Cells already ending in `"EAST"` are ignored.

---

## 📝 Notes

- Cells with errors are **ignored**.
- Highlight is cleared from all column F cells before each run.
- Match is **case-insensitive** using regular expressions (e.g., `ast`, `Ast`, etc.).
- Useful for verifying corrections from macros like `ReplaceASTWithEast`.

---

## ✍️ Author & Info

- **Author:** Ilesanmi Kehinde John (aka Calm)  
- **Created On:** July 29, 2025  
- **Purpose:** Validation step to check for uncaught address suffixes still ending in `AST`  
- **License:** Open for personal or team use

---
